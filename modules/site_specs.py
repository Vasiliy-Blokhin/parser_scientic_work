"""
Спецификации для подключения к различным сайтам и сервисам
с открытой возможностью скачивания (минимальная регистрация или без неё).

Каждая спецификация описывает:
- name:         имя сервиса
- base_url:     корневой URL
- search_path:  путь поиска (шаблон с {query} и {page})
- needs_auth:   требуется ли регистрация
- pdf_url_tmpl: шаблон URL для скачивания PDF
- link_selector: CSS-селектор ссылок на статьи
- article_domain: домен статей (фильтр)
- article_path_contains: проверка пути на принадлежность к статье
"""

SITE_SPECS = [
    {
        "name": "CyberLeninka",
        "base_url": "https://cyberleninka.ru",
        "search_path": "/search?q={query}&page={page}",
        "needs_auth": False,
        "pdf_url_tmpl": "{article_url}/pdf",
        "link_selector": 'a[href*="/article/"]',
        "article_domain": "cyberleninka.ru",
        "article_path_contains": "/article/",
        "use_browser_download": True,
    },
    # --- Заготовки для будущих сервисов ---
    # eLibrary — требует регистрации, но API доступен
    # {
    #     "name": "eLibrary",
    #     "base_url": "https://elibrary.ru",
    #     "search_path": "/search.asp?searchquery={query}&page={page}",
    #     "needs_auth": True,
    #     "auth_type": "cookie",
    #     "pdf_url_tmpl": None,  # Скачивание через интерфейс
    #     "link_selector": 'a[href*="/item.asp"]',
    #     "article_domain": "elibrary.ru",
    #     "article_path_contains": "/item.asp",
    #     "use_browser_download": False,
    # },
    # ResearchGate — открытые PDF без регистрации
    # {
    #     "name": "ResearchGate",
    #     "base_url": "https://www.researchgate.net",
    #     "search_path": "/search?q={query}&page={page}",
    #     "needs_auth": False,
    #     "pdf_url_tmpl": None,
    #     "link_selector": 'a[href*="/publication/"]',
    #     "article_domain": "www.researchgate.net",
    #     "article_path_contains": "/publication/",
    #     "use_browser_download": True,
    # },
]


def get_site_spec(name):
    """Возвращает спецификацию по имени сервиса."""
    for spec in SITE_SPECS:
        if spec["name"].lower() == name.lower():
            return spec
    return None


def get_default_site():
    """Возвращает спецификацию сервиса по умолчанию."""
    return SITE_SPECS[0] if SITE_SPECS else None
