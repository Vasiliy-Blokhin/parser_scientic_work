"""
Парсер научных статей с CyberLeninka.

Рефакторинг по ТЗ (версия 2):
  1. Контроль дубликатов на начальном этапе (MD5) + спецификации
  2. Извлечение информации на этапе парсинга + спецификация документа
  3. Формирование директорий с блоками по 10 документов
  4. Пул ключевых слов с приоритетами (техническая литература — первой)
  5. Спецификации для подключения к различным сайтам и сервисам
  6. Разделение main на этапы с логированием и комментариями

Обновление (версия 3):
  7. Сохранение информации о скачанных файлах в SQLite-БД
  8. Предварительная проверка дубликатов через БД (отсекает до скачивания)
  9. Спецификации хранятся рядом с PDF-файлами в блочных директориях
  10. Циклический поиск по пулу ключевых слов (блоками по 10 документов)
  11. Работа до остановки пользователем или достижения 1 ГБ суммарного веса
"""

import hashlib
import json
import logging
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
)
from modules.values import SHORT_ABBREVIATIONS, STOP_WORDS
from modules.keyword_pool import get_keywords_by_priority, get_keyword_groups
from modules.site_specs import get_default_site, get_site_spec

# Обновлённые модули (с интеграцией БД)
from modules.db_manager import DatabaseManager
from modules.spec_manager import SpecManager


# ──────────────────────────────────────────────────────────────
#  Константы
# ──────────────────────────────────────────────────────────────

# Лимит суммарного веса скачанных файлов (1 ГБ в байтах)
MAX_TOTAL_SIZE_BYTES = 1 * 1024 * 1024 * 1024


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
    Парсер научных статей.

    Использование:
        parser = Parser()
        parser()                      # интерактивный режим
        parser(use_pool=True)         # режим с пулом ключевых слов (без лимита, до 1 ГБ)
        parser(use_pool=True, max_results=50)  # режим пула с лимитом по количеству
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
        self._block_number = None

        # Счётчики для циклического режима
        self._total_downloaded_size = 0
        self._total_articles = 0

    # ════════════════════════════════════════════════════════════
    #  TOC: __call__ — точка входа
    # ════════════════════════════════════════════════════════════

    def __call__(self, *args, **kwargs):
        """
        Точка входа. Разделена на этапы с логированием.
        """
        use_pool = kwargs.get("use_pool", False)
        max_results = kwargs.get("max_results", None)  # None = без лимита (до 1 ГБ)

        # ── Этап 1: Подготовка ──
        logger.info("=" * 60)
        logger.info("ЗАПУСК ПАРСЕРА (v3 — БД + циклический режим)")
        logger.info("=" * 60)
        self._stage_prepare()

        # ── Этап 2: Сбор настроек ──
        settings = self._stage_settings(use_pool=use_pool, max_results=max_results)

        # ── Этап 3: Запуск браузера ──
        browser_ctx = self._stage_browser_start()
        if browser_ctx is None:
            logger.error("Не удалось запустить браузер. Остановка.")
            return

        # ── Этап 4: Парсинг ──
        try:
            articles = self._stage_parse(settings, browser_ctx)

            # ── Этап 5: Сохранение результатов ──
            self._stage_save(articles, settings)

            # ── Этап 6: Статистика ──
            self._stage_stats()

        except KeyboardInterrupt:
            logger.info("Остановлено пользователем.")
            self._stage_stats()
        except Exception as e:
            logger.error("КРИТИЧЕСКАЯ ОШИБКА: %s", repr(e))
        finally:
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

        # Текущий блок
        self._block_number = self.spec_manager.current_block
        logger.info("Текущий блок директорий: %d", self._block_number)

    # ════════════════════════════════════════════════════════════
    #  ЭТАП 2: Настройки
    # ════════════════════════════════════════════════════════════

    def _stage_settings(self, use_pool=False, max_results=None):
        """Сбор настроек поиска."""
        logger.info("Этап 2: Настройки поиска")

        if use_pool:
            keywords = get_keywords_by_priority()
            logger.info("Пул ключевых слов (%d шт.): %s",
                        len(keywords), keywords[:5])

            settings = {
                "queries": keywords,
                "start_year": 2020,
                "end_year": 2026,
                "max_results": max_results,  # None = без лимита
                "max_pages": 5,
                "filters": [],
                "filter_mode": "any",
                "download_pdf": True,
                "max_total_size": MAX_TOTAL_SIZE_BYTES,
            }
            if max_results is None:
                logger.info(
                    "Режим пула: %d запросов, без лимита статей, лимит %.0f GB",
                    len(settings["queries"]),
                    settings["max_total_size"] / (1024**3),
                )
            else:
                logger.info(
                    "Режим пула: %d запросов, лимит статей %d, лимит %.0f GB",
                    len(settings["queries"]),
                    max_results,
                    settings["max_total_size"] / (1024**3),
                )
        else:
            settings = self._get_settings_interactive()
            settings["max_total_size"] = MAX_TOTAL_SIZE_BYTES

        logger.info(
            "Настройки: запросов=%d, годы=%d-%d, лимит=%s, PDF=%s, размер-лимит=%.2f GB",
            len(settings["queries"]),
            settings["start_year"],
            settings["end_year"],
            settings["max_results"],
            settings["download_pdf"],
            settings["max_total_size"] / (1024**3),
        )
        return settings

    @staticmethod
    def _get_settings_interactive():
        """Интерактивный ввод настроек."""
        print()
        print("=" * 70)
        print("НАСТРОЙКИ ПОИСКА")
        print("=" * 70)
        print()
        print("Введите поисковые слова или фразы.")
        print("Можно использовать ',' или ';' как разделитель.")
        print()
        print("Пример:")
        print("ИИ, искусственный интеллект, RAG, машинное обучение")

        query = input("Поиск: ").strip()

        queries = [
            x.strip()
            for x in re.split(r"[;,]", query)
            if x.strip()
        ]

        if not queries:
            queries = ["нейросети"]

        start_year = int(input("Начальный год [2024]: ") or 2024)
        end_year = int(input("Конечный год [2026]: ") or 2026)

        max_results_input = input("Максимум статей [пусто = без лимита, до 1 ГБ]: ").strip()
        max_results = int(max_results_input) if max_results_input else None

        max_pages = int(input("Максимум страниц на запрос [5]: ") or 5)

        print()
        print("Дополнительный фильтр. Можно оставить пустым.")
        print("Разделители: ',' или ';'")

        filter_text = input("Фильтр: ").strip()
        filters = [
            x.strip()
            for x in re.split(r"[;,]", filter_text)
            if x.strip()
        ]

        filter_mode = "any"
        if filters:
            print()
            print("1 — достаточно любого")
            print("2 — нужны все")
            if input("Режим [1]: ").strip() == "2":
                filter_mode = "all"

        print()
        download_pdf_flag = (
            input("Скачивать статьи в PDF? [Y/n]: ").strip().lower() != "n"
        )

        return {
            "queries": queries,
            "start_year": start_year,
            "end_year": end_year,
            "max_results": max_results,
            "max_pages": max_pages,
            "filters": filters,
            "filter_mode": filter_mode,
            "download_pdf": download_pdf_flag,
        }

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
                extra_http_headers={
                    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"
                },
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
    #  ЭТАП 4: Парсинг — основной цикл (циклический режим)
    # ════════════════════════════════════════════════════════════

    def _stage_parse(self, settings, browser_ctx):
        """
        Основной цикл парсинга (циклический режим).

        Проходит по пулу ключевых слов циклически, постранично.
        Каждый цикл: для каждого ключевого слова — по 1 странице.
        Продолжает, пока не достигнут лимит (1 ГБ или max_results)
        или пользователь не остановит (Ctrl+C).
        """
        logger.info("Этап 4: Парсинг (циклический режим)")

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

        articles = []
        seen_urls = set()

        # Загружаем URL, уже записанные в БД, чтобы не дублировать
        db_urls = self.db_manager.get_known_urls()
        seen_urls.update(db_urls)
        logger.info("URL из БД: %d (исключаются из поиска)", len(db_urls))

        queries = settings["queries"]
        max_pages = settings["max_pages"]
        max_results = settings["max_results"]
        max_total_size = settings["max_total_size"]

        # ── Циклический проход по пулу ключевых слов ──
        cycle_number = 0
        no_new_articles_cycles = 0
        MAX_IDLE_CYCLES = 3  # если 3 цикла подряд нет новых статей — стоп

        while True:
            # Проверка лимитов
            if max_results is not None and len(articles) >= max_results:
                logger.info("Достигнут лимит по количеству: %d статей", max_results)
                break

            if self._total_downloaded_size >= max_total_size:
                logger.info(
                    "Достигнут лимит по размеру: %.2f GB",
                    self._total_downloaded_size / (1024**3),
                )
                break

            cycle_number += 1
            new_this_cycle = 0
            logger.info("=" * 50)
            logger.info("ЦИКЛ %d | Статей: %d | Размер: %.2f MB / %.2f GB",
                        cycle_number, len(articles),
                        self._total_downloaded_size / (1024**2),
                        max_total_size / (1024**3))
            logger.info("=" * 50)

            for query in queries:
                # Проверка лимитов внутри цикла
                if max_results is not None and len(articles) >= max_results:
                    break
                if self._total_downloaded_size >= max_total_size:
                    break

                logger.info("─" * 40)
                logger.info("Ключевое слово: %s", query)
                logger.info("─" * 40)

                # Для каждого ключевого слова — по одной странице за цикл
                # Номер страницы = cycle_number (ограничено max_pages)
                page_number = cycle_number
                if page_number > max_pages:
                    page_number = ((cycle_number - 1) % max_pages) + 1

                search_url = self._build_search_url(query, page_number)
                logger.info("Страница поиска %d", page_number)

                html = self.get_html(page, search_url)
                if not html:
                    logger.warning("Поиск не загружен, пропуск")
                    continue

                links = self.search_links(html)
                logger.info("Ссылок найдено: %d", len(links))

                for article_url in links:
                    # Проверка лимитов
                    if max_results is not None and len(articles) >= max_results:
                        break
                    if self._total_downloaded_size >= max_total_size:
                        break

                    # ── Предварительная проверка через БД ──
                    is_dup, existing = self.db_manager.check_duplicate_before_download(article_url)
                    if is_dup:
                        logger.info("✗ Дубликат по URL в БД, пропуск: %s", article_url)
                        continue

                    if article_url in seen_urls:
                        continue
                    seen_urls.add(article_url)

                    logger.info("Кандидат #%d: %s", len(seen_urls), article_url)

                    article = self._process_article(
                        page, article_url, query, settings
                    )

                    if article:
                        articles.append(article)
                        new_this_cycle += 1
                        logger.info(
                            "✓ Статья #%d. Размер: %.2f MB / %.2f GB",
                            len(articles),
                            self._total_downloaded_size / (1024**2),
                            max_total_size / (1024**3),
                        )
                        time.sleep(ARTICLE_DELAY)

                time.sleep(SEARCH_DELAY)

            # Проверка: были ли новые статьи в этом цикле
            if new_this_cycle == 0:
                no_new_articles_cycles += 1
                logger.warning(
                    "Цикл %d: новых статей нет (%d/%d)",
                    cycle_number, no_new_articles_cycles, MAX_IDLE_CYCLES,
                )
                if no_new_articles_cycles >= MAX_IDLE_CYCLES:
                    logger.info(
                        "Прекращение: %d циклов без новых статей.",
                        MAX_IDLE_CYCLES,
                    )
                    break
            else:
                no_new_articles_cycles = 0

        logger.info("Парсинг завершён. Найдено статей: %d", len(articles))
        return articles

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

    def _process_article(self, page, article_url, query, settings):
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
            logger.info("✗ Не подходит по году")
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
            return None

        # ── 6. Проверка дубликата через БД (по URL — ещё раз, на случай гонки) ──
        is_dup, _ = self.db_manager.check_duplicate_before_download(article_url)
        if is_dup:
            logger.info("✗ Дубликат по URL в БД, пропуск")
            article["Статус PDF"] = "duplicate"
            return None

        # ── 7. Скачивание PDF ──
        if settings["download_pdf"]:
            # Ротация блока (10 документов → новый блок)
            self._check_block_rotation()

            pdf_path, pdf_status, pdf_link = self.download_pdf(
                page, article_url, article_title, article_year
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
                )
                if is_dup:
                    logger.info("✗ Дубликат по MD5, пропуск")
                    article["Статус PDF"] = "duplicate"
                    return None
            else:
                logger.warning("PDF не скачан, спецификация не создана")
                article["Статус PDF"] = "error"
        else:
            article["Статус PDF"] = "disabled"

        return article

    # ════════════════════════════════════════════════════════════
    #  Управление блоками директорий
    # ════════════════════════════════════════════════════════════

    def _check_block_rotation(self):
        """Проверяет, не заполнен ли текущий блок (10 документов)."""
        if self.spec_manager.maybe_rotate_block():
            self._block_number = self.spec_manager.current_block
            logger.info("Новый блок директорий: block_%03d", self._block_number)

    def _get_current_block_dir(self):
        """Возвращает директорию текущего блока."""
        return self.spec_manager.get_block_dir(self.spec_manager.current_block)

    def _check_and_save(self, pdf_path, title, authors, abstract,
                        url, year, doi, journal, keywords, query):
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
                block_number=self.spec_manager.current_block,
                query_keyword=query,
                status="duplicate",
            )
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
            block_number=self.spec_manager.current_block,
            query_keyword=query,
            status="downloaded",
        )

        # Обновляем суммарный размер
        if file_size:
            self._total_downloaded_size += file_size
            self._total_articles += 1

        # ── 3. Создание спецификации (рядом с PDF, в той же директории) ──
        spec_path = self.spec_manager.create_spec(
            file_path=pdf_path,
            title=title,
            authors=authors,
            abstract=abstract,
            url=url,
            year=year,
            doi=doi,
            keywords=keywords,
            block_number=self.spec_manager.current_block,
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
        """Сохранение результатов в Excel."""
        logger.info("Этап 5: Сохранение результатов")

        if not articles:
            logger.info("Подходящих статей нет.")
            return

        self.save_excel(
            articles,
            settings["queries"],
            settings["start_year"],
            settings["end_year"],
        )

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
                logger.info("  %s: %d файлов, %.2f MB", s["keyword"], s["count"], s["size_mb"])

        # Статистика по блокам
        block_stats = self.db_manager.get_stats_by_block()
        if block_stats:
            logger.info("── По блокам ──")
            for s in block_stats:
                logger.info("  block_%03d: %d файлов, %.2f MB", s["block"], s["count"], s["size_mb"])

        # Статистика из spec_manager
        spec_stats = self.spec_manager.get_stats()
        logger.info("── Файловая система ──")
        logger.info("Всего PDF: %d", spec_stats["total_docs"])
        logger.info("Размер: %.2f MB", spec_stats["total_size_mb"])
        logger.info("Блоков: %d", spec_stats["blocks"])
        if spec_stats.get("block_numbers"):
            logger.info(
                "Номера блоков: %s",
                ", ".join(f"block_{n:03d}" for n in spec_stats["block_numbers"]),
            )

    # ════════════════════════════════════════════════════════════
    #  Утилиты
    # ════════════════════════════════════════════════════════════

    @staticmethod
    def check_directory():
        """Создание базовых директорий."""
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        PDF_DIR.mkdir(parents=True, exist_ok=True)

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
    def cache_file(url):
        """Один HTML-файл на URL."""
        filename = hashlib.sha1(url.encode("utf-8")).hexdigest() + ".html"
        return CACHE_DIR / filename

    def get_cached(self, url):
        """Читает HTML из кэша."""
        if not USE_CACHE:
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

            if "/search" not in url or "/article/" in html:
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
    def safe_pdf_name(title, year, article_url):
        """Безопасное имя PDF-файла."""
        title = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", title)
        title = re.sub(r"\s+", " ", title).strip()
        title = title[:100].rstrip()
        short_hash = hashlib.sha1(article_url.encode("utf-8")).hexdigest()[:8]
        return f"{year} - {title} - {short_hash}.pdf"

    def download_pdf(self, page, article_url, title, year):
        """
        Скачивание PDF через браузерный download.
        Сохраняет в директорию текущего блока (рядом со спецификацией).
        """
        pdf_url = article_url.rstrip("/") + "/pdf"
        filename = self.safe_pdf_name(title, year, article_url)

        block_dir = self._get_current_block_dir()
        block_dir.mkdir(parents=True, exist_ok=True)
        path = block_dir / filename

        # Файл уже существует
        if SKIP_EXISTING_PDFS and path.exists() and path.stat().st_size > 0:
            logger.info("PDF уже существует: %s", path)
            return str(path), "already_exists", pdf_url

        for attempt in range(1, PDF_RETRIES + 1):
            logger.info("PDF попытка %d/%d: %s", attempt, PDF_RETRIES, pdf_url)
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
                logger.warning("PDF ошибка (попытка %d): %s", attempt, repr(e))
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
    def save_excel(articles, queries, start_year, end_year):
        """Сохранение результатов в Excel."""
        if not articles:
            return

        df = pd.DataFrame(articles)
        df = df.drop_duplicates(subset="Ссылка")

        if "DOI" in df.columns:
            df["_doi"] = df["DOI"].fillna("").astype(str).str.lower().str.strip()
            with_doi = df[df["_doi"] != ""].drop_duplicates(subset="_doi")
            without_doi = df[df["_doi"] == ""]
            df = pd.concat([with_doi, without_doi], ignore_index=True)
            df = df.drop(columns=["_doi"])

        df = df.sort_values(["Релевантность", "Год"], ascending=[False, False])

        query_name = re.sub(r"[^а-яА-Яa-zA-Z0-9_-]+", "_", "_".join(queries))
        query_name = query_name[:60]
        filename = f"cyberleninka_{query_name}_{start_year}_{end_year}.xlsx"

        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Статьи")
            ws = writer.sheets["Статьи"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            widths = {
                "A": 55, "B": 30, "C": 10, "D": 30, "E": 25,
                "F": 25, "G": 25, "H": 15, "I": 35, "J": 30,
                "K": 50, "L": 70, "M": 70, "N": 70, "O": 20,
            }
            for column, width in widths.items():
                ws.column_dimensions[column].width = width

            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = cell.alignment.copy(
                        wrap_text=True, vertical="top"
                    )

            headers = {
                cell.value: cell.column for cell in ws[1] if cell.value
            }

            if "Ссылка" in headers:
                column = headers["Ссылка"]
                for row_num in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row_num, column=column)
                    if cell.value:
                        cell.hyperlink = cell.value
                        cell.style = "Hyperlink"

            if "PDF URL" in headers:
                column = headers["PDF URL"]
                for row_num in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row_num, column=column)
                    if cell.value:
                        cell.hyperlink = cell.value
                        cell.style = "Hyperlink"

            if "PDF" in headers:
                column = headers["PDF"]
                for row_num in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row_num, column=column)
                    if cell.value:
                        pdf_path = Path(str(cell.value))
                        if pdf_path.exists():
                            cell.hyperlink = pdf_path.resolve().as_uri()
                            cell.style = "Hyperlink"

        logger.info("Excel сохранён: %s", filename)


# ──────────────────────────────────────────────────────────────
#  Точка входа
# ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = Parser()

    # Интерактивный режим:
    # parser()

    # Режим с пулом ключевых слов (работает до 1 ГБ или Ctrl+C):
    parser(use_pool=True)

    # Режим с пулом и лимитом по количеству:
    # parser(use_pool=True, max_results=50)
