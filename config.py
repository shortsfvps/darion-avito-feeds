"""
Файл глобальной конфигурации генератора XML-фидов Авито компании «Дарион Свет».

Здесь задаются пути к исходным файлам, параметры синхронизации с Google Таблицами,
а также настройки публикации фидов на удаленный веб-сервер по протоколам FTP / SFTP.
"""

import os

# Базовые пути проекта
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

# Путь к номенклатурному файлу каталога
NOMENCLATURE_FILE = os.path.join(DATA_DIR, "Продукты (19).xlsx")

# Ссылка на актуальные остатки и цены (Google Sheets экспорт в формате Excel)
GOOGLE_SHEETS_URL = "https://docs.google.com/spreadsheets/d/1ooUhgWojjoce0YGbzrmzJtL9tz9xm_OU0QBSaqEMCas/export?format=xlsx"

# Локальные файлы кэша остатков
LATEST_STOCK_FILE = os.path.join(DATA_DIR, "latest_stock.xlsx")
TEMPLATE_STOCK_FILE = os.path.join(DATA_DIR, "remains_spb_template.xlsx")
TAGS_SCHEMA_FILE = os.path.join(DATA_DIR, "template_tags_schema.json")

# Таймаут запроса к Google Таблице (в секундах)
GOOGLE_SHEETS_TIMEOUT = 15

# ==============================================================================
# НАСТРОЙКИ ПУБЛИКАЦИИ ФИДОВ (FTP / SFTP)
# ==============================================================================
# Варианты UPLOAD_METHOD:
# - 'local' : файлы сохраняются только локально в папке output/
# - 'ftp'   : после генерации файлы автоматически загружаются на FTP-сервер
# - 'sftp'  : после генерации файлы автоматически передаются по защищенному SFTP (SSH)
UPLOAD_METHOD = 'local'

# Настройки для FTP
FTP_HOST = ''
FTP_PORT = 21
FTP_USER = ''
FTP_PASS = ''
FTP_REMOTE_DIR = '/public_html/avito_feeds'
FTP_PUBLIC_URL_BASE = 'http://darion-svet.com/avito_feeds'

# Настройки для SFTP (SSH)
SFTP_HOST = ''
SFTP_PORT = 22
SFTP_USER = ''
SFTP_PASS = ''
SFTP_REMOTE_DIR = '/var/www/darion-svet/feeds'
SFTP_PUBLIC_URL_BASE = 'http://darion-svet.com/avito_feeds'

# ==============================================================================
# НАСТРОЙКИ ОРГАНИЗАЦИИ И ОБЪЯВЛЕНИЙ АВИТО
# ==============================================================================
COMPANY_NAME = "Дарион Свет"
COMPANY_BRAND = "VIRONA"
DEFAULT_ADDRESS = "Санкт-Петербург, Студенческая ул., 10"

# Конфигурация 8 целевых фидов Авито
# Схема категорий Авито XML v.3:
# Раздел сайта «Для дома и дачи» в XML не используется.
# Верхний тег <Category>: «Мебель и интерьер» (для светотехники) или «Ремонт и строительство» (для электрики).
# Подкатегория <GoodsType>: «Освещение» (для светотехники) или «Электрика» (для электрики).
# Подтип <GoodsSubType>: «Потолочные светильники», «Уличное освещение», «Лампочки», «Светильники» и т.д.
FEEDS_CONFIG = {
    "led_luminaires": {
        "filename": "led_luminaires_feed.xml",
        "title": "LED светильники (промышленные, потолочные, настенные)",
        "category": "Мебель и интерьер",
        "goods_type": "Освещение",
        "goods_sub_type": "Потолочное и настенное",
        "lighting_type": "Люстры и потолочные светильники",
        "template": "led_luminaires.xml"
    },
    "track_systems": {
        "filename": "track_systems_feed.xml",
        "title": "Уличное освещение (консольные, уличные светильники)",
        "category": "Мебель и интерьер",
        "goods_type": "Освещение",
        "goods_sub_type": "Уличное",
        "template": "track_systems.xml"
    },
    "chandeliers": {
        "filename": "chandeliers_feed.xml",
        "title": "Люстры и потолочные светильники",
        "category": "Мебель и интерьер",
        "goods_type": "Освещение",
        "goods_sub_type": "Люстры и потолочные светильники",
        "template": "chandeliers.xml"
    },
    "lamps": {
        "filename": "lamps_feed.xml",
        "title": "Лампы и лампочки",
        "category": "Мебель и интерьер",
        "goods_type": "Освещение",
        "goods_sub_type": "Комплектующие",
        "lighting_type": "Лампочки",
        "template": "lamps.xml"
    },
    "lighting_fixtures": {
        "filename": "lighting_fixtures_feed.xml",
        "title": "Светильники общего назначения (Бра, Споты)",
        "category": "Мебель и интерьер",
        "goods_type": "Освещение",
        "goods_sub_type": "Светильники",
        "template": "lighting_fixtures.xml"
    },
    "lamp_luminaires": {
        "filename": "lamp_luminaires_feed.xml",
        "title": "Ламповые светильники",
        "category": "Мебель и интерьер",
        "goods_type": "Освещение",
        "goods_sub_type": "Светильники",
        "template": "lamp_luminaires.xml"
    },
    "electrics": {
        "filename": "electrics_feed.xml",
        "title": "Электрика и комплектующие",
        "category": "Ремонт и строительство",
        "goods_type": "Электрика",
        "goods_sub_type": "Комплектующие и материалы",
        "template": "electrics.xml"
    },
    "other_goods": {
        "filename": "other_goods_feed.xml",
        "title": "Прочие товары / Торговое оборудование",
        "category": "Мебель и интерьер",
        "goods_type": "Освещение",
        "goods_sub_type": "Другое",
        "template": "other_goods.xml"
    }
}
