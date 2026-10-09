"""
Парсер научных статей с CyberLeninka.

Версия 4 — непрерывный режим:
  1. Только пул ключевых слов (без интерактивного режима)
  2. Непрерывное скачивание, остановка только по Ctrl+C
  3. Случайная задержка 10–50 секунд между запросами
  4. Динамическое обновление данных из БД в начале каждого цикла
  5. Категории → блоки по 10 документов
  6. Пул исключений, фильтр по годам (2024–2026)
  7. Быстрый перезапуск: обработанные URL и прогресс хранятся в БД
"""

import hashlib
import json
import logging
import random
import re
import time
import urllib.parse
from pathlib import Path

import pandas as pd

try:
    import pymorphy3
except Exception:
    pymorphy3 = None

from bs4 import BeautifulSoup
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)

from modules.settings import (
    BASE_URL, BROWSER, HEADLESS, USE_PROXY, PROXY, TIMEOUT, PDF_TIMEOUT,
    USE_CACHE, CACHE_DIR, PDF_DIR, DOWNLOAD_PDF, PDF_RETRIES,
    SKIP_EXISTING_PDFS, SEARCH_DELAY, ARTICLE_DELAY, MIN_SCORE,
    MAX_ABSTRACT_LENGTH, MIN_SCORE_ONE_WORD,
    MIN_REQUEST_DELAY, MAX_REQUEST_DELAY,
    START_YEAR, END_YEAR, MAX_PAGES, FILTERS, FILTER_MODE,
    USE_EXCLUSIONS, EXCLUSION_MIN_ABSTRACT_HITS, DOCS_PER_BLOCK,
)
from modules.values import SHORT_ABBREVIATIONS, STOP_WORDS
from modules.keyword_pool import (
    get_keywords_by_priority, get_keyword_groups, get_query_category_map,
)
from modules.exclusion_pool import get_all_exclusion_phrases
from modules.site_specs import get_default_site, get_site_spec

# Обновлённые модули (с интеграцией БД)
from modules.db_manager import DatabaseManager
from modules.spec_manager import SpecManager


# ──────────────────────────────────────────────────────────────
#  Логирование
# ──────────────────────────────────────────────────────────────

logger = logging.getLogger("parser")
logger.setLevel(logging.DEBUG)

_console = logging.StreamHandler()
_console.setLevel(logging.INFO)
_console.setFormatter(
    logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
)
logger.addHandler(_console)

_file = logging.FileHandler("parser.log", encoding="utf-8")
_file.setLevel(logging.DEBUG)
_file.setFormatter(
    logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
)
logger.addHandler(_file)


class Parser:
    """
    Парсер научных статей (непрерывный режим).

    Использование:
        parser = Parser()
        parser()
    """

    def __init__(self):
        if pymorphy3 is not None:
            try:
                self.morph_analyzer = pymorphy3.MorphAnalyzer()
            except Exception:
                self.morph_analyzer = None
                logger.warning("pymorphy3 не инициализирован — работа без лемматизации")
        else:
            self.morph_analyzer = None

        # БД для учёта скачанных файлов и проверки дубликатов
        self.db_manager = DatabaseManager()

        # Менеджер спецификаций (с интеграцией БД)
        self.spec_manager = SpecManager(db_manager=self.db_manager)

        self.site_spec = get_default_site()

        # Категория по ключевому слову + фразы исключений
        self.query_category = get_query_category_map()
        self.exclusion_patterns = self._build_exclusion_patterns()

        # Был ли реальный сетевой запрос (для умной задержки)
        self._network_used = False

        # Счётчики
        self._total_downloaded_size = 0
        self._total_articles = 0
        self._articles = []

    # ════════════════════════════════════════════════════════════
    #  TOC: __call__ — точка входа
    # ════════════════════════════════════════════════════════════

    def __call__(self, *args, **kwargs):
        """
        Точка входа. Непрерывный режим — только пул ключевых слов.
        Остановка только по Ctrl+C.
        """
        logger.info("=" * 60)
        logger.info("ЗАПУСК ПАРСЕРА (v4 — непрерывный режим)")
        logger.info("Задержка между запросами: %d–%d сек",
                    MIN_REQUEST_DELAY, MAX_REQUEST_DELAY)
        logger.info("Размер блока: 10 документов")
        logger.info("=" * 60)

        # ── Этап 1: Подготовка ──
        self._stage_prepare()

        # ── Этап 2: Настройки ──
        settings = self._stage_settings()

        # ── Этап 3: Запуск браузера ──
        browser_ctx = self._stage_browser_start()
        if browser_ctx is None:
            logger.error("Не удалось запустить браузер. Остановка.")
            return

        # ── Этап 4: Парсинг (бесконечный цикл) ──
        try:
            self._stage_parse(settings, browser_ctx)
        except KeyboardInterrupt:
            logger.info("Остановлено пользователем.")
        except Exception as e:
            logger.error("КРИТИЧЕСКАЯ ОШИБКА: %s", repr(e))
        finally:
            # ── Этап 5: Сохранение результатов ──
            if self._articles:
                self._stage_save(self._articles, settings)

            # ── Этап 6: Статистика ──
            self._stage_stats()

            # ── Этап 7: Завершение ──
            self._stage_browser_stop(browser_ctx)
            self.db_manager.close()

    # ════════════════════════════════════════════════════════════
    #  ЭТАП 1: Подготовка
    # ════════════════════════════════════════════════════════════

    def _stage_prepare(self):
        """Создание директорий, инициализация БД и менеджера спецификаций."""
        logger.info("Этап 1: Подготовка директорий и БД")

        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        PDF_DIR.mkdir(parents=True, exist_ok=True)
        self.spec_manager.init_dirs()

        # Загружаем известные хеши из БД
        known = len(self.spec_manager.known_hashes)
        logger.info("Известных документов (по MD5 из БД): %d", known)

        # Текущий суммарный размер из БД
        self._total_downloaded_size = self.db_manager.get_total_size()
        self._total_articles = self.db_manager.get_downloaded_count()
        logger.info(
            "Уже скачано: %d файлов, %.2f MB (%.4f GB)",
            self._total_articles,
            self._total_downloaded_size / (1024 * 1024),
            self._total_downloaded_size / (1024 * 1024 * 1024),
        )

        logger.info(
            "Обработано URL (скачано + отклонено): %d",
            len(self.db_manager.get_processed_urls()),
        )

    # ════════════════════════════════════════════════════════════
    #  ЭТАП 2: Настройки (только пул ключевых слов)
    # ════════════════════════════════════════════════════════════

    def _stage_settings(self):
        """Сбор настроек поиска — только пул ключевых слов."""
        logger.info("Этап 2: Настройки поиска (режим пула ключевых слов)")

        keywords = get_keywords_by_priority()
        logger.info("Пул ключевых слов (%d шт.): %s",
                    len(keywords), keywords[:5])

        settings = {
            "queries": keywords,
            "start_year": START_YEAR,
            "end_year": END_YEAR,
            "max_pages": MAX_PAGES,
            "filters": FILTERS,
            "filter_mode": FILTER_MODE,
            "download_pdf": True,
        }
        logger.info(
            "Настройки: запросов=%d, годы=%d-%d, PDF=%s",
            len(settings["queries"]),
            settings["start_year"],
            settings["end_year"],
            settings["download_pdf"],
        )
        return settings

    # ════════════════════════════════════════════════════════════
    #  ЭТАП 3: Запуск браузера
    # ════════════════════════════════════════════════════════════

    def _stage_browser_start(self):
        """Запускает браузер и возвращает (playwright, browser, context, page)."""
        logger.info("Этап 3: Запуск браузера (%s)", BROWSER)

        playwright = None
        browser = None
        context = None
        page = None

        try:
            playwright = sync_playwright().start()

            if BROWSER == "firefox":
                engine = playwright.firefox
            else:
                engine = playwright.chromium

            options = {"headless": HEADLESS}
            if USE_PROXY:
                options["proxy"] = {"server": PROXY}

            # Stealth: отключаем признак автоматизации
            if BROWSER == "chromium":
                options["args"] = [
                    "--disable-blink-features=AutomationControlled",
                ]

            try:
                browser = engine.launch(**options)
            except Exception as e:
                playwright.stop()
                if "Executable doesn't exist" in str(e):
                    raise RuntimeError(
                        f"\nБраузер {BROWSER} не установлен.\n\n"
                        f"Выполните:\n"
                        f"python -m playwright install {BROWSER}\n"
                    ) from e
                raise

            context = browser.new_context(
                locale="ru-RU",
                viewport={"width": 1400, "height": 900},
                accept_downloads=True,
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
                extra_http_headers={
                    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"
                },
            )

            # Скрываем navigator.webdriver
            context.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', "
                "{get: () => undefined})"
            )

            page = context.new_page()
            page.set_default_timeout(TIMEOUT)
            page.route("**/*", self._route_handler)

            logger.info(
                "Браузер запущен. Прокси: %s",
                PROXY if USE_PROXY else "отключен",
            )
            return (playwright, browser, context, page)

        except Exception as e:
            logger.error("Ошибка запуска браузера: %s", repr(e))
            if playwright:
                playwright.stop()
            return None

    @staticmethod
    def _stage_browser_stop(browser_ctx):
        """Корректно закрывает браузер."""
        if browser_ctx is None:
            return

        playwright, browser, context, page = browser_ctx
        logger.info("Этап 7: Завершение работы браузера")

        for name, obj, action in [
            ("context", context, "close"),
            ("browser", browser, "close"),
            ("playwright", playwright, "stop"),
        ]:
            try:
                if obj:
                    getattr(obj, action)()
                    logger.debug("Закрыт: %s", name)
            except Exception as e:
                logger.warning("Ошибка закрытия %s: %s", name, e)

        logger.info("Работа завершена.")

    # ════════════════════════════════════════════════════════════
    #  ЭТАП 4: Парсинг — непрерывный цикл
    # ════════════════════════════════════════════════════════════

    def _random_delay(self):
        """Случайная задержка между запросами (10–50 сек)."""
        delay = random.uniform(MIN_REQUEST_DELAY, MAX_REQUEST_DELAY)
        logger.info("Задержка: %.1f сек", delay)
        time.sleep(delay)

    def _stage_parse(self, settings, browser_ctx):
        """
        Непрерывный цикл парсинга.

        Проходит по пулу ключевых слов циклически, постранично.
        Каждый цикл: для каждого ключевого слова — по 1 странице.
        Работает бесконечно, остановка только по Ctrl+C.

        В начале каждого цикла — динамическое обновление данных из БД.
        Между запросами — случайная задержка 10–50 секунд.
        """
        logger.info("Этап 4: Непрерывный парсинг")

        _, _, _, page = browser_ctx

        # Проверка доступности сайта
        site = self.site_spec
        response = page.goto(
            site["base_url"],
            wait_until="domcontentloaded",
            timeout=TIMEOUT,
        )

        if not response or response.status >= 400:
            raise RuntimeError(
                f"Сайт {site['name']} недоступен (статус: "
                f"{response.status if response else 'нет ответа'})"
            )

        logger.info("%s доступен: статус %d", site["name"], response.status)

        queries = settings["queries"]
        max_pages = settings["max_pages"]

        # ── Восстановление прогресса прошлого запуска ──
        cycle_number = self.db_manager.get_state("cycle", 0) or 1
        done_queries = set(self.db_manager.get_state("done_queries", []))
        if cycle_number > 1 or done_queries:
            logger.info(
                "Возобновление: цикл %d, уже пройдено запросов: %d",
                cycle_number, len(done_queries),
            )

        while True:
            # ── Динамическое обновление данных из БД ──
            # processed = скачанные + отклонённые (год, релевантность,
            # исключения, дубликаты). Их повторно не открываем.
            processed = self.db_manager.get_processed_urls()
            self._total_downloaded_size = self.db_manager.get_total_size()
            self._total_articles = self.db_manager.get_downloaded_count()
            new_this_cycle = 0

            self.db_manager.set_state("cycle", cycle_number)

            logger.info("=" * 50)
            logger.info("ЦИКЛ %d | Скачано: %d | Размер: %.2f MB | Обработано URL: %d",
                        cycle_number,
                        self._total_articles,
                        self._total_downloaded_size / (1024 * 1024),
                        len(processed))
            logger.info("=" * 50)

            page_number = ((cycle_number - 1) % max_pages) + 1

            for query in queries:
                if query in done_queries:
                    logger.info("Пропуск (уже пройден в этом цикле): %s", query)
                    continue

                category = self.query_category.get(query, "Без категории")
                logger.info("─" * 40)
                logger.info("[%s] Ключевое слово: %s | страница %d",
                            category, query, page_number)
                logger.info("─" * 40)

                search_url = self._build_search_url(query, page_number)

                self._network_used = False
                html = self.get_html(page, search_url)
                if self._network_used:
                    self._random_delay()
                if not html:
                    logger.warning("Поиск не загружен, пропуск")
                    continue

                links = self.search_links(html)
                new_links = [u for u in links if u not in processed]
                logger.info("Ссылок: %d, новых: %d", len(links), len(new_links))

                for article_url in new_links:
                    processed.add(article_url)
                    logger.info("Кандидат: %s", article_url)

                    self._network_used = False
                    article = self._process_article(
                        page, article_url, query, category, settings
                    )

                    if article:
                        self._articles.append(article)
                        new_this_cycle += 1
                        logger.info(
                            "✓ Статья #%d [%s / block_%03d]. Размер: %.2f MB",
                            len(self._articles), category, article["Блок"],
                            self._total_downloaded_size / (1024 * 1024),
                        )

                    # Задержка только если был реальный запрос к сайту
                    if self._network_used:
                        self._random_delay()

                # Запрос пройден — запоминаем прогресс
                done_queries.add(query)
                self.db_manager.set_state("done_queries", sorted(done_queries))

            # ── Сохранение Excel (из БД: все запуски, не только этот) ──
            logger.info("Сохранение Excel...")
            self.save_excel(self.db_manager.get_all_records())

            logger.info(
                "Цикл %d завершён. Новых статей: %d. Всего в БД: %d",
                cycle_number, new_this_cycle,
                self.db_manager.get_downloaded_count(),
            )

            # Следующий цикл
            cycle_number += 1
            done_queries = set()
            self.db_manager.set_state("cycle", cycle_number)
            self.db_manager.set_state("done_queries", [])
            self._random_delay()

    def _build_search_url(self, query, page_number):
        """Формирует URL поиска по спецификации сайта."""
        search_path = self.site_spec["search_path"]
        return (
            self.site_spec["base_url"]
            + search_path.format(
                query=urllib.parse.quote(query),
                page=page_number,
            )
        )

    def _process_article(self, page, article_url, query, category, settings):
        """
        Обработка одной статьи:
        1. Загрузка HTML
        2. Извлечение метаданных
        3. Проверка года
        4. Оценка релевантности
        5. Дополнительный фильтр
        6. Проверка дубликата через БД (предварительная, по URL)
        7. Скачивание PDF
        8. Контроль дубликата по MD5 (после скачивания)
        9. Сохранение в БД + создание спецификации (рядом с PDF)
        """
        # ── 1. Загрузка HTML ──
        article_html = self.get_html(page, article_url)
        if not article_html:
            logger.warning("HTML не загружен: %s", article_url)
            # error не блокирует повторную попытку при следующем запуске
            self.db_manager.mark_processed(article_url, "error", "no_html")
            return None

        soup = BeautifulSoup(article_html, "html.parser")

        # ── 2. Извлечение метаданных ──
        article_title = self.get_title(soup)
        article_year = self.get_year(soup)
        article_keywords = self.get_keywords(soup)
        article_abstract = self.get_abstract(soup)
        article_authors = self.get_authors(soup)
        article_doi = self.get_doi(soup)
        article_journal = self.get_journal(soup)

        logger.info("Название: %s", article_title[:160])
        logger.info("Год: %s", article_year)

        # ── 3. Проверка года ──
        if (
            article_year is None
            or not (settings["start_year"] <= article_year <= settings["end_year"])
        ):
            logger.info("✗ Не подходит по году (%s, нужно %d–%d)",
                        article_year, settings["start_year"],
                        settings["end_year"])
            self.db_manager.mark_processed(
                article_url, "rejected", f"year:{article_year}"
            )
            return None

        # ── 3.1 Пул исключений ──
        if USE_EXCLUSIONS:
            excluded = self.check_exclusions(
                article_title, article_keywords, article_abstract
            )
            if excluded:
                logger.info("✗ Исключена пулом: %s", excluded)
                self.db_manager.mark_processed(
                    article_url, "excluded", excluded
                )
                return None

        # ── 4. Оценка релевантности ──
        article_score, matches = self.score_article(
            query, article_title, article_keywords, article_abstract
        )

        logger.info("Score: %d, совпадения: %s", article_score, matches)

        threshold = (
            MIN_SCORE_ONE_WORD
            if len(self.words(query)) == 1
            else MIN_SCORE
        )

        if article_score < threshold:
            logger.info("✗ Низкая релевантность (%d < %d)", article_score, threshold)
            self.db_manager.mark_processed(
                article_url, "rejected", f"score:{article_score}"
            )
            return None

        # Формируем запись статьи
        article = {
            "Название": article_title,
            "Авторы": article_authors,
            "Год": article_year,
            "Журнал": article_journal,
            "DOI": article_doi,
            "Релевантность": article_score,
            "Совпадения": matches,
            "Категория": category,
            "Блок": 0,
            "Поисковый запрос": query,
            "Ключевые слова": article_keywords,
            "Аннотация": article_abstract,
            "Ссылка": article_url,
            "PDF": "",
            "PDF URL": article_url.rstrip("/") + "/pdf",
            "Статус PDF": "",
        }

        # ── 5. Дополнительный фильтр ──
        if not self.passes_filter(
            article, settings["filters"], settings["filter_mode"]
        ):
            logger.info("✗ Не прошла дополнительный фильтр")
            self.db_manager.mark_processed(article_url, "rejected", "filter")
            return None

        # ── 6. Проверка дубликата через БД (по URL — ещё раз, на случай гонки) ──
        is_dup, _ = self.db_manager.check_duplicate_before_download(article_url)
        if is_dup:
            logger.info("✗ Дубликат по URL в БД, пропуск")
            article["Статус PDF"] = "duplicate"
            self.db_manager.mark_processed(article_url, "duplicate", "url")
            return None

        # ── 7. Скачивание PDF ──
        if settings["download_pdf"]:
            # Ротация блока (10 документов → новый блок)
            self._check_block_rotation(category)
            article["Блок"] = self.spec_manager.get_current_block(category)

            self._network_used = True
            pdf_path, pdf_status, pdf_link = self.download_pdf(
                page, article_url, article_title, article_year, category
            )

            article["PDF"] = pdf_path
            article["PDF URL"] = pdf_link
            article["Статус PDF"] = pdf_status

            # ── 8. Контроль дубликата по MD5 + БД + спецификация ──
            if pdf_path and Path(pdf_path).exists():
                is_dup = self._check_and_save(
                    pdf_path=pdf_path,
                    title=article_title,
                    authors=article_authors,
                    abstract=article_abstract,
                    url=article_url,
                    year=article_year,
                    doi=article_doi,
                    journal=article_journal,
                    keywords=article_keywords,
                    query=query,
                    category=category,
                    score=article_score,
                    matches=matches,
                    pdf_url=pdf_link,
                )
                if is_dup:
                    logger.info("✗ Дубликат по MD5, пропуск")
                    article["Статус PDF"] = "duplicate"
                    self.db_manager.mark_processed(
                        article_url, "duplicate", "md5"
                    )
                    return None
                self.db_manager.mark_processed(article_url, "downloaded")
            else:
                logger.warning("PDF не скачан, спецификация не создана")
                article["Статус PDF"] = "error"
                self.db_manager.mark_processed(article_url, "error", "pdf")
        else:
            article["Статус PDF"] = "disabled"
            self.db_manager.mark_processed(article_url, "downloaded", "no_pdf")

        return article

    # ════════════════════════════════════════════════════════════
    #  Управление блоками директорий
    # ════════════════════════════════════════════════════════════

    def _check_block_rotation(self, category):
        """Проверяет, не заполнен ли текущий блок категории."""
        if self.spec_manager.maybe_rotate_block(category):
            logger.info(
                "[%s] Новый блок: block_%03d",
                category, self.spec_manager.get_current_block(category),
            )

    def _get_current_block_dir(self, category):
        """Директория текущего блока категории."""
        return self.spec_manager.get_block_dir(
            category, self.spec_manager.get_current_block(category)
        )

    # ──────────────────────────────────────────────────────────
    #  Пул исключений
    # ──────────────────────────────────────────────────────────

    def _build_exclusion_patterns(self):
        """Компилирует фразы исключений в регексы (совпадение по началу слов)."""
        patterns = []
        for phrase, group in get_all_exclusion_phrases():
            norm = self.normalize(phrase)
            if not norm:
                continue
            parts = [re.escape(w) for w in norm.split()]
            regex = re.compile(r"\b" + r"\w*\s+".join(parts))
            patterns.append((regex, phrase, group))
        return patterns

    def check_exclusions(self, title, keywords, abstract):
        """
        Возвращает строку-причину, если статья попадает под исключения,
        иначе "".

        Название / ключевые слова: достаточно одной фразы.
        Аннотация: нужно EXCLUSION_MIN_ABSTRACT_HITS разных фраз
        (чтобы случайное упоминание не отсекало техническую статью).
        """
        head = self.normalize(f"{title} {keywords}")
        for regex, phrase, group in self.exclusion_patterns:
            if regex.search(head):
                return f"{group}: «{phrase}» (название/ключевые слова)"

        body = self.normalize(abstract)
        hits = [
            (phrase, group)
            for regex, phrase, group in self.exclusion_patterns
            if regex.search(body)
        ]
        if len({ph for ph, _ in hits}) >= EXCLUSION_MIN_ABSTRACT_HITS:
            names = ", ".join(sorted({ph for ph, _ in hits})[:5])
            return f"{hits[0][1]}: {names} (аннотация)"
        return ""

    def _check_and_save(self, pdf_path, title, authors, abstract,
                        url, year, doi, journal, keywords, query,
                        category="", score=None, matches="", pdf_url=""):
        """
        Полный цикл сохранения:
        1. Проверка дубликата по MD5 (через spec_manager)
        2. Сохранение записи в БД
        3. Создание спецификации (рядом с PDF)

        Возвращает True, если найден дубликат.
        """
        pdf_path = Path(pdf_path)

        # ── 1. Проверка дубликата по MD5 ──
        if self.spec_manager.is_duplicate(pdf_path):
            logger.warning("Дубликат по MD5: %s", pdf_path.name)

            # Сохраняем запись о дубликате в БД
            self.db_manager.save_file_record(
                url=url,
                file_path=pdf_path,
                title=title,
                authors=authors,
                year=year,
                journal=journal,
                doi=doi,
                keywords=keywords,
                abstract=abstract,
                block_number=self.spec_manager.get_current_block(category),
                query_keyword=query,
                status="duplicate",
                category=category,
                score=score,
                matches=matches,
                pdf_url=pdf_url,
            )
            # Копия не нужна — иначе занимает место в блоке
            try:
                pdf_path.unlink()
            except OSError:
                pass
            return True

        # ── 2. Сохранение записи в БД ──
        md5_hash, file_size = self.db_manager.save_file_record(
            url=url,
            file_path=pdf_path,
            title=title,
            authors=authors,
            year=year,
            journal=journal,
            doi=doi,
            keywords=keywords,
            abstract=abstract,
            block_number=self.spec_manager.get_current_block(category),
            query_keyword=query,
            status="downloaded",
            category=category,
            score=score,
            matches=matches,
            pdf_url=pdf_url,
        )

        # Обновляем суммарный размер
        if file_size:
            self._total_downloaded_size += file_size
            self._total_articles += 1

        # ── 3. Создание спецификации (рядом с PDF, в той же директории) ──
        spec_path = self.spec_manager.create_spec(
            file_path=pdf_path,
            category=category,
            title=title,
            authors=authors,
            abstract=abstract,
            url=url,
            year=year,
            doi=doi,
            journal=journal,
            keywords=keywords,
            block_number=self.spec_manager.get_current_block(category),
        )

        if spec_path:
            logger.info("Спецификация создана рядом с PDF: %s", spec_path.name)
        else:
            logger.error("Не удалось создать спецификацию")

        return False

    # ════════════════════════════════════════════════════════════
    #  ЭТАП 5: Сохранение результатов
    # ════════════════════════════════════════════════════════════

    def _stage_save(self, articles, settings):
        """Сохранение результатов в Excel (все статьи из БД)."""
        logger.info("Этап 5: Сохранение результатов")
        self.save_excel(self.db_manager.get_all_records())

    # ════════════════════════════════════════════════════════════
    #  ЭТАП 6: Статистика
    # ════════════════════════════════════════════════════════════

    def _stage_stats(self):
        """Выводит статистику по БД, спецификациям и дубликатам."""
        logger.info("Этап 6: Статистика")

        # Статистика из БД
        db_stats = self.db_manager.get_stats()
        logger.info("── БД ──")
        logger.info("Скачанных файлов: %d", db_stats["total_files"])
        logger.info("Дубликатов: %d", db_stats["duplicates"])
        logger.info("Суммарный размер: %.2f MB (%.4f GB)",
                    db_stats["total_size_mb"], db_stats["total_size_gb"])

        # Статистика по ключевым словам
        kw_stats = self.db_manager.get_stats_by_keyword()
        if kw_stats:
            logger.info("── По ключевым словам ──")
            for s in kw_stats[:10]:
                logger.info("  %s: %d файлов, %.2f MB",
                            s["keyword"], s["count"], s["size_mb"])

        # Статистика из spec_manager: категории → блоки
        spec_stats = self.spec_manager.get_stats()
        logger.info("── Файловая система ──")
        logger.info("Всего PDF: %d (%.2f MB)",
                    spec_stats["total_docs"], spec_stats["total_size_mb"])
        for name, st in spec_stats["categories"].items():
            logger.info("  %s: %d PDF, блоков %d, %.2f MB",
                        name, st["docs"], st["blocks"], st["size_mb"])

    # ════════════════════════════════════════════════════════════
    #  Утилиты
    # ════════════════════════════════════════════════════════════

    @staticmethod
    def check_directory():
        """Создание базовых директорий."""
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        PDF_DIR.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def cache_file(url):
        """Один HTML-файл на URL."""
        filename = hashlib.sha1(url.encode("utf-8")).hexdigest() + ".html"
        return CACHE_DIR / filename

    def get_cached(self, url):
        """Читает HTML из кэша."""
        if not USE_CACHE:
            return None
        # Не кэшируем поисковые страницы — они могут меняться
        if "/search" in url:
            return None
        path = self.cache_file(url)
        if not path.exists():
            return None
        try:
            html = path.read_text(encoding="utf-8")
            if html.strip():
                logger.debug("Кэшhit: %s", url)
                return html
        except Exception as e:
            logger.warning("Ошибка чтения кэша %s: %s", url, e)
        return None

    def save_cached(self, url, html):
        """Сохраняет HTML в кэш."""
        if not USE_CACHE or not html:
            return
        # Не кэшируем поисковые страницы
        if "/search" in url:
            return
        try:
            self.cache_file(url).write_text(html, encoding="utf-8")
            logger.debug("Кэш сохранён: %s", url)
        except Exception as e:
            logger.warning("Ошибка записи кэша %s: %s", url, e)

    @staticmethod
    def _route_handler(route):
        """Блокирует загрузку изображений, шрифтов и медиа."""
        if route.request.resource_type in {"image", "font", "media"}:
            route.abort()
        else:
            route.continue_()

    def get_html(self, page, url):
        """Загружает страницу через Playwright."""
        cached = self.get_cached(url)
        if cached:
            return cached

        logger.debug("GET: %s", url)
        self._network_used = True

        try:
            response = page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=TIMEOUT,
            )
            if not response:
                return None

            logger.debug("STATUS: %d", response.status)
            if response.status >= 400:
                return None

            if "/search" in url:
                try:
                    page.wait_for_selector(
                        'a[href*="/article/"]', timeout=15_000
                    )
                except PlaywrightTimeoutError:
                    page.wait_for_timeout(1500)
            else:
                page.wait_for_timeout(300)

            html = page.content()
            if not html:
                return None

            self.save_cached(url, html)

            logger.debug("HTML: %d bytes", len(html.encode("utf-8")))
            return html

        except Exception as e:
            logger.warning("Ошибка загрузки %s: %s", url, repr(e))
            return None

    @staticmethod
    def meta(soup, *names):
        """Получает содержимое meta-тега."""
        for name in names:
            tag = soup.find("meta", attrs={"name": name})
            if not tag:
                tag = soup.find("meta", attrs={"property": name})
            if tag and tag.get("content"):
                return tag["content"].strip()
        return ""

    def get_title(self, soup):
        value = self.meta(soup, "citation_title")
        if value:
            return value
        value = self.meta(soup, "og:title")
        if value:
            return re.sub(
                r"\s*[–—-]\s*тема научной статьи.*$", "", value,
                flags=re.IGNORECASE,
            ).strip()
        h1 = soup.find("h1")
        if h1:
            value = h1.get_text(" ", strip=True)
            if value:
                return value
        return "Без названия"

    def get_year(self, soup):
        value = self.meta(
            soup, "citation_publication_date", "citation_date",
            "article:published_time", "date", "DC.date", "DC.Date",
        )
        match = re.search(r"\b(20\d{2})\b", value)
        if match:
            return int(match.group(1))

        for tag in soup.select(
            '[itemprop="datePublished"],'
            '[itemprop="dateCreated"],'
            "time"
        ):
            value = tag.get("datetime") or tag.get_text(" ", strip=True)
            match = re.search(r"\b(20\d{2})\b", value)
            if match:
                return int(match.group(1))

        for script in soup.find_all("script", type="application/ld+json"):
            text = script.get_text(" ", strip=True)
            match = re.search(
                r'"datePublished"\s*:\s*"[^"]*(20\d{2})',
                text, flags=re.IGNORECASE,
            )
            if match:
                return int(match.group(1))
        return None

    @staticmethod
    def get_authors(soup):
        authors = []
        for tag in soup.find_all("meta", attrs={"name": "citation_author"}):
            value = (tag.get("content") or "").strip()
            if value:
                authors.append(value)
        if not authors:
            for element in soup.select('[itemprop="author"]'):
                name = element.select_one('[itemprop="name"]')
                if name:
                    value = name.get_text(" ", strip=True)
                else:
                    value = element.get_text(" ", strip=True)
                if value:
                    authors.append(value)
        return "; ".join(dict.fromkeys(authors)) or "Не указаны"

    def get_keywords(self, soup):
        value = self.meta(soup, "citation_keywords", "keywords")
        if value:
            return value

        text = soup.get_text(" ", strip=True)
        match = re.search(
            r"(?:ключевые\s+слова|keywords)\s*:?\s*(.{0,1000})",
            text, flags=re.IGNORECASE,
        )
        if not match:
            return ""
        value = re.split(
            r"\b(?:аннотация|abstract|введение|introduction)\b",
            match.group(1), maxsplit=1, flags=re.IGNORECASE,
        )[0]
        return value.strip()

    def get_abstract(self, soup):
        value = self.meta(
            soup, "citation_abstract", "description", "og:description",
        )
        if value:
            return value[:MAX_ABSTRACT_LENGTH]

        text = soup.get_text(" ", strip=True)
        match = re.search(
            r"(?:аннотация|abstract)\s*:?\s*(.{50,5000})",
            text, flags=re.IGNORECASE,
        )
        if not match:
            return ""
        return match.group(1).strip()[:MAX_ABSTRACT_LENGTH]

    def get_doi(self, soup):
        value = self.meta(soup, "citation_doi")
        if not value:
            value = soup.get_text(" ", strip=True)
        match = re.search(
            r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+",
            value, flags=re.IGNORECASE,
        )
        if not match:
            return ""
        return match.group(0).rstrip(".,;)")

    def get_journal(self, soup):
        return self.meta(soup, "citation_journal_title")

    @staticmethod
    def search_links(html):
        """Извлекает ссылки на статьи из HTML."""
        soup = BeautifulSoup(html, "html.parser")
        result = []
        seen = set()

        for tag in soup.select('a[href*="/article/"]'):
            href = tag.get("href", "").split("#")[0]
            if not href:
                continue
            url = urllib.parse.urljoin(BASE_URL, href)
            parsed = urllib.parse.urlparse(url)

            if parsed.netloc not in {"cyberleninka.ru", "www.cyberleninka.ru"}:
                continue
            if "/article/" not in parsed.path:
                continue

            url = urllib.parse.urlunparse(
                (parsed.scheme, parsed.netloc, parsed.path, "", "", "")
            )
            if url in seen:
                continue
            seen.add(url)
            result.append(url)

        return result

    def score_article(self, query, title, keywords, abstract):
        """Рассчитывает релевантность статьи."""
        query_normalized = self.normalize(query)
        query_words = set(self.words(query))
        title_words = set(self.words(title))
        keyword_words = set(self.words(keywords))
        abstract_words = set(self.words(abstract))

        score = 0
        matches = []

        for word in query_words:
            if word in title_words:
                score += 10
                value = f"title:{word}"
                if value not in matches:
                    matches.append(value)
            if word in keyword_words:
                score += 7
                value = f"keywords:{word}"
                if value not in matches:
                    matches.append(value)
            if word in abstract_words:
                score += 3
                value = f"abstract:{word}"
                if value not in matches:
                    matches.append(value)

        found = query_words & (title_words | keyword_words | abstract_words)
        if query_words:
            coverage = len(found) / len(query_words)
            if coverage == 1:
                score += 15
                value = "title:phrase"
                if value not in matches:
                    matches.insert(0, value)
            elif coverage >= 0.75:
                score += 8

        return score, "; ".join(matches)

    def passes_filter(self, article, filters, mode):
        """Дополнительный фильтр по содержимому статьи."""
        if not filters:
            return True

        text = " ".join([
            article["Название"], article["Авторы"],
            article["Ключевые слова"], article["Аннотация"],
        ])
        text_words = set(self.words(text))

        checks = []
        for item in filters:
            required = set(self.words(item))
            if required:
                checks.append(required <= text_words)

        if not checks:
            return True

        if mode == "all":
            return all(checks)
        return any(checks)

    @staticmethod
    def normalize(text):
        """Нормализация текста."""
        text = str(text or "").lower()
        text = text.replace("ё", "е")
        text = re.sub(r"[^а-яa-z0-9\s-]", " ", text)
        return re.sub(r"\s+", " ", text).strip()

    def words(self, text):
        """Получает список лемм."""
        result = []
        for original in str(text or "").split():
            original = original.strip()
            if not original:
                continue
            normalized = self.normalize(original)
            if not normalized:
                continue
            is_abbreviation = (normalized in SHORT_ABBREVIATIONS)
            if len(normalized) < 3 and not is_abbreviation:
                continue
            if normalized in STOP_WORDS:
                continue
            if normalized not in result:
                result.append(normalized)
        return result

    @staticmethod
    def safe_pdf_name(title, year, article_url):
        """Безопасное имя PDF-файла."""
        title = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", title)
        title = re.sub(r"\s+", " ", title).strip()
        title = title[:100].rstrip()
        short_hash = hashlib.sha1(article_url.encode("utf-8")).hexdigest()[:8]
        return f"{year} - {title} - {short_hash}.pdf"

    def download_pdf(self, page, article_url, title, year, category=""):
        """
        Скачивание PDF через браузерный download.
        Сохраняет в директорию текущего блока (рядом со спецификацией).
        """
        pdf_url = article_url.rstrip("/") + "/pdf"
        filename = self.safe_pdf_name(title, year, article_url)

        block_dir = self._get_current_block_dir(category)
        block_dir.mkdir(parents=True, exist_ok=True)
        path = block_dir / filename

        # Файл уже существует
        if SKIP_EXISTING_PDFS and path.exists() and path.stat().st_size > 0:
            logger.info("PDF уже существует: %s", path)
            return str(path), "already_exists", pdf_url

        for attempt in range(1, PDF_RETRIES + 1):
            logger.info("Скачивание PDF (попытка %d/%d): %s",
                        attempt, PDF_RETRIES, pdf_url)
            pdf_page = None

            try:
                pdf_page = page.context.new_page()
                pdf_page.set_default_timeout(PDF_TIMEOUT)

                with pdf_page.expect_download(timeout=PDF_TIMEOUT) as download_info:
                    try:
                        pdf_page.goto(
                            pdf_url,
                            wait_until="commit",
                            timeout=PDF_TIMEOUT,
                        )
                    except Exception as e:
                        if "Download is starting" not in str(e):
                            raise

                download = download_info.value
                download.save_as(str(path))

                if not path.exists() or path.stat().st_size <= 0:
                    raise RuntimeError("PDF не был сохранён")

                size = path.stat().st_size
                logger.info(
                    "PDF сохранён: %s (%.2f MB)", path, size / 1024 / 1024
                )
                return str(path), "downloaded", pdf_url

            except Exception as e:
                logger.warning("Ошибка скачивания (попытка %d): %s",
                               attempt, repr(e))
                if attempt < PDF_RETRIES:
                    time.sleep(attempt * 2)
            finally:
                if pdf_page:
                    try:
                        pdf_page.close()
                    except Exception:
                        pass

        return "", "error", pdf_url

    @staticmethod
    def _sheet_name(name, used):
        """Допустимое и уникальное имя листа Excel (≤31 символ)."""
        base = re.sub(r"[\\/*?:\[\]]", "_", name)[:31] or "Лист"
        result, i = base, 2
        while result in used:
            suffix = f" {i}"
            result = base[:31 - len(suffix)] + suffix
            i += 1
        used.add(result)
        return result

    @staticmethod
    def save_excel(records):
        """
        Сохраняет Excel: отдельный лист на каждую категорию,
        внутри листа — блоки (по DOCS_PER_BLOCK статей) с заголовками.
        records — строки из БД (db_manager.get_all_records()).
        """
        from openpyxl.styles import Font, PatternFill, Alignment

        if not records:
            return

        columns = [
            ("Блок", "block_number", 8),
            ("Название", "title", 55),
            ("Авторы", "authors", 30),
            ("Год", "year", 8),
            ("Журнал", "journal", 25),
            ("DOI", "doi", 25),
            ("Релевантность", "score", 14),
            ("Совпадения", "matches", 30),
            ("Поисковый запрос", "query_keyword", 30),
            ("Ключевые слова", "keywords", 40),
            ("Аннотация", "abstract", 70),
            ("Ссылка", "url", 45),
            ("PDF URL", "pdf_url", 45),
            ("PDF", "file_path", 40),
        ]

        by_category = {}
        for r in records:
            by_category.setdefault(r.get("category") or "Без категории", []).append(r)

        filename = "cyberleninka_results.xlsx"
        block_fill = PatternFill("solid", fgColor="DDEBF7")
        head_fill = PatternFill("solid", fgColor="BDD7EE")
        used_names = set()

        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            # Лист-сводка
            summary = []
            for cat, rows in by_category.items():
                blocks = {r.get("block_number") for r in rows}
                summary.append({
                    "Категория": cat,
                    "Статей": len(rows),
                    "Блоков": len(blocks),
                })
            pd.DataFrame(summary).to_excel(
                writer, index=False, sheet_name="Сводка"
            )
            used_names.add("Сводка")

            for cat, rows in by_category.items():
                name = Parser._sheet_name(cat, used_names)
                ws = writer.book.create_sheet(name)
                writer.sheets[name] = ws

                # Шапка таблицы
                for col, (title, _, width) in enumerate(columns, start=1):
                    cell = ws.cell(row=1, column=col, value=title)
                    cell.font = Font(bold=True)
                    cell.fill = head_fill
                    ws.column_dimensions[cell.column_letter].width = width
                ws.freeze_panes = "A2"

                blocks = {}
                for r in rows:
                    blocks.setdefault(r.get("block_number") or 0, []).append(r)

                row_num = 2
                for block_no in sorted(blocks):
                    items = sorted(
                        blocks[block_no],
                        key=lambda r: (-(r.get("score") or 0), -(r.get("year") or 0)),
                    )
                    # Строка-заголовок блока
                    ws.cell(
                        row=row_num, column=1,
                        value=f"Блок {block_no:03d} — {len(items)} из {DOCS_PER_BLOCK} статей",
                    )
                    for col in range(1, len(columns) + 1):
                        c = ws.cell(row=row_num, column=col)
                        c.fill = block_fill
                        c.font = Font(bold=True)
                    row_num += 1

                    for r in items:
                        for col, (_, key, _) in enumerate(columns, start=1):
                            value = r.get(key)
                            if key == "block_number":
                                value = block_no
                            cell = ws.cell(row=row_num, column=col, value=value)
                            cell.alignment = Alignment(
                                wrap_text=True, vertical="top"
                            )
                            if key in ("url", "pdf_url") and value:
                                cell.hyperlink = value
                                cell.style = "Hyperlink"
                            elif key == "file_path" and value:
                                pdf_path = Path(str(value))
                                if pdf_path.exists():
                                    cell.hyperlink = pdf_path.resolve().as_uri()
                                    cell.style = "Hyperlink"
                        row_num += 1
                    row_num += 1  # пустая строка между блоками

        logger.info("Excel сохранён: %s (категорий: %d)", filename, len(by_category))


# ──────────────────────────────────────────────────────────────
#  Точка входа
# ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = Parser()
    parser()
