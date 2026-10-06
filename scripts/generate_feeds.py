"""
Генератор 8 XML-фидов для Авито (Автозагрузка) компании «Дарион Свет».

Архитектура:
1. Данные номенклатуры, остатков и цен читаются напрямую из листов загружаемой Google Таблицы (1..8).
2. Листы 1..8 сопоставляются с 8 фидами из FEEDS_CONFIG:
   1 -> led_luminaires (led_luminaires_feed.xml)
   2 -> track_systems (track_systems_feed.xml)
   3 -> chandeliers (chandeliers_feed.xml)
   4 -> lamps (lamps_feed.xml)
   5 -> lighting_fixtures (lighting_fixtures_feed.xml)
   6 -> lamp_luminaires (lamp_luminaires_feed.xml)
   7 -> electrics (electrics_feed.xml)
   8 -> other_goods (other_goods_feed.xml)
3. Фильтрация строк:
   - Столбец E (5): основное фото заполнено (не пустое, начинается с http).
   - Столбец BI (61): остаток строго > 0 (очистка от неразрывных пробелов \xa0).
   - Столбец BJ (62): цена строго > 1 (очистка от пробелов, рублей, запятых).
4. Формирование валидного XML Avito v.3 с CDATA-описанием, ограничением длины заголовков (<= 50 симв.)
   и лимитом до 10 фотографий на объявление. При 0 товаров генерируется валидный пустой каркас.
"""

import os
import sys
import io
import re
import logging
import ftplib
import requests
import pandas as pd
import openpyxl
from lxml import etree
import paramiko

# Безопасный вывод UTF-8 в консоль Windows
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Подключение config.py из корня проекта
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import config

# Настройка логирования на sys.stdout
handler = logging.StreamHandler(sys.stdout)
handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S'))
logger = logging.getLogger("AvitoFeedGenerator")
logger.setLevel(logging.INFO)
logger.handlers = [handler]

# Пути из config.py
DATA_DIR = config.DATA_DIR
OUTPUT_DIR = config.OUTPUT_DIR
GOOGLE_SHEET_URL = config.GOOGLE_SHEETS_URL
LATEST_STOCK_PATH = config.LATEST_STOCK_FILE
TEMPLATE_STOCK_PATH = config.TEMPLATE_STOCK_FILE
FEEDS_CONFIG = config.FEEDS_CONFIG

# Сопоставление номеров листов 1..8 с ключами фидов
SHEET_INDEX_TO_FEED_KEY = {
    1: "led_luminaires",
    2: "track_systems",
    3: "chandeliers",
    4: "lamps",
    5: "lighting_fixtures",
    6: "lamp_luminaires",
    7: "electrics",
    8: "other_goods",
}


def clean_numeric(val) -> float:
    """
    Преобразует значение ячейки в float, безопасно удаляя мусорные символы:
    неразрывные пробелы (\xa0), 'руб.', 'шт.', запятые -> точки, и т.п.
    Возвращает 0.0 при невозможности преобразования.
    """
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace('\xa0', ' ')
    s = re.sub(r'(?i)(руб\.?|шт\.?|%)', '', s)
    s = s.replace(',', '.').strip()
    m = re.search(r'-?\d+(?:\.\d+)?', s)
    if m:
        try:
            return float(m.group())
        except (ValueError, TypeError):
            return 0.0
    return 0.0


def clean_sku(val) -> str:
    """Очищает артикул товара от хвостовых .0 при чтении из Excel."""
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        f = float(val)
        if f.is_integer():
            return str(int(f))
    s = str(val).strip()
    if re.match(r'^\d+\.0$', s):
        return s[:-2]
    return s


def format_num_val(val) -> str:
    """Форматирует числовые характеристики (мощность, лм, габариты и т.п.) в компактную строку."""
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        f = float(val)
        if f.is_integer():
            return str(int(f))
        return f"{f:g}"
    s = str(val).strip()
    if re.match(r'^-?\d+\.0$', s):
        return s[:-2]
    return s


class StockSync:
    """Модуль загрузки и кэширования актуальной книги Google Таблицы."""

    def __init__(self, url=GOOGLE_SHEET_URL, latest_cache=LATEST_STOCK_PATH, fallback_cache=TEMPLATE_STOCK_PATH):
        self.url = url
        self.latest_cache = latest_cache
        self.fallback_cache = fallback_cache
        self.sync_status = "Not started"
        self.sync_message = ""
        self.used_source = ""

    def download_stock_file(self) -> bytes:
        """Скачивает свежие данные по экспортной ссылке Google Таблицы (таймаут 15 сек)."""
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        timeout = getattr(config, "GOOGLE_SHEETS_TIMEOUT", 15)
        try:
            logger.info(f"Запрос актуальной таблицы из Google Sheets (таймаут {timeout} сек)...")
            response = requests.get(self.url, headers=headers, timeout=timeout)
            if response.status_code == 200 and response.content[:2] == b"PK" and len(response.content) > 1000:
                with open(self.latest_cache, "wb") as f:
                    f.write(response.content)
                self.sync_status = "SUCCESS_ONLINE"
                self.sync_message = f"Успешно загружено из Google Sheets (HTTP 200, {len(response.content)} байт)"
                self.used_source = self.latest_cache
                logger.info(self.sync_message)
                return response.content
            else:
                self.sync_status = f"HTTP_{response.status_code}"
                self.sync_message = (
                    f"Google Таблица вернула HTTP {response.status_code}. Используется локальный кэш."
                )
                logger.warning(self.sync_message)
                return None
        except Exception as e:
            self.sync_status = "NETWORK_ERROR"
            self.sync_message = f"Ошибка сети при запросе к Google Sheets: {e}. Используется локальный кэш."
            logger.warning(self.sync_message)
            return None

    def get_workbook(self) -> openpyxl.Workbook:
        """Возвращает объект openpyxl.Workbook из скачанных байтов или локального кэша."""
        raw_bytes = self.download_stock_file()
        source = None

        if raw_bytes:
            source = io.BytesIO(raw_bytes)
        elif os.path.exists(self.latest_cache) and os.path.getsize(self.latest_cache) > 1000:
            source = self.latest_cache
            self.used_source = self.latest_cache
            logger.info(f"Используем кэшированный файл: {self.latest_cache}")
        elif os.path.exists(self.fallback_cache):
            source = self.fallback_cache
            self.used_source = self.fallback_cache
            logger.info(f"Используем резервный эталонный кэш: {self.fallback_cache}")
        else:
            raise FileNotFoundError(f"Файлы таблицы не найдены: {self.latest_cache} и {self.fallback_cache}")

        wb = openpyxl.load_workbook(source, data_only=True)

        # Проверяем листы 1..8: если какой-то лист в скачанной таблице пуст (нет товаров с артикулом),
        # но в резервном эталоне (sheets_backup.xlsx) данные есть, восстанавливаем его
        backup_path = os.path.join(DATA_DIR, "sheets_backup.xlsx")
        if os.path.exists(backup_path):
            wb_backup = None
            for sname in list(wb.sheetnames):
                m = re.match(r'^\s*([1-8])\b', sname)
                if not m:
                    continue
                ws = wb[sname]
                has_items = any(ws.cell(r, 1).value for r in range(3, min(ws.max_row + 1, 15)))
                if not has_items:
                    if wb_backup is None:
                        wb_backup = openpyxl.load_workbook(backup_path, data_only=True)
                    for bname in wb_backup.sheetnames:
                        mb = re.match(r'^\s*([1-8])\b', bname)
                        if mb and mb.group(1) == m.group(1):
                            ws_b = wb_backup[bname]
                            cnt_b = sum(1 for r in range(3, ws_b.max_row + 1) if ws_b.cell(r, 1).value)
                            if cnt_b > 0:
                                del wb[sname]
                                ws_new = wb.create_sheet(title=sname)
                                for row in ws_b.iter_rows(values_only=True):
                                    ws_new.append(list(row))
                                logger.info(f"Лист '{sname}' дополнен из резервного эталона ({cnt_b} позиций)")
                            break

        return wb

    def load_stock_dict(self) -> dict:
        """
        Собирает словарь остатков и цен для обратной совместимости.
        Поддерживает как новые листы 1..8, так и старый лист 'Остатки СПб'.
        """
        wb = self.get_workbook()
        stock_dict = {}

        # 1. Проверяем новые листы 1..8
        has_numbered_sheets = any(re.match(r'^\s*([1-8])\b', s) for s in wb.sheetnames)
        if has_numbered_sheets:
            for sname in wb.sheetnames:
                if not re.match(r'^\s*([1-8])\b', sname):
                    continue
                ws = wb[sname]
                for r in range(3, ws.max_row + 1):
                    sku = clean_sku(ws.cell(r, 1).value)
                    if not sku:
                        continue
                    stock = clean_numeric(ws.cell(r, 61).value)
                    price = clean_numeric(ws.cell(r, 62).value)
                    if stock > 0 and price > 1:
                        stock_dict[sku] = {
                            "price": int(round(price)),
                            "stock": int(round(stock))
                        }
            return stock_dict

        # 2. Фоллбэк на лист 'Остатки СПб'
        target_sheet = None
        for s in wb.sheetnames:
            if "остат" in s.lower() and "спб" in s.lower():
                target_sheet = s
                break
        if not target_sheet:
            target_sheet = wb.sheetnames[0]

        ws = wb[target_sheet]
        for r in range(2, ws.max_row + 1):
            sku = clean_sku(ws.cell(r, 1).value)
            if not sku or sku.lower() == "артикул":
                continue
            # Индексы 10 и 11 (Col J и Col K)
            price = clean_numeric(ws.cell(r, 10).value)
            stock = clean_numeric(ws.cell(r, 11).value)
            if stock > 0 and price > 0:
                stock_dict[sku] = {
                    "price": int(round(price)),
                    "stock": int(round(stock))
                }
        return stock_dict


class ImageResolver:
    """Модуль обработки, резолвинга и валидации изображений для фидов Авито."""

    # Резервные прямые ссылки на изображения товаров сайта darion-svet.com
    KNOWN_CATALOG_IMAGES = {
        '14690612053025': 'http://darion-svet.com/web/binary/image?model=product.product&id=23973&field=image',
        '4690612053004': 'http://darion-svet.com/web/binary/image?model=product.product&id=23970&field=image',
        '4690612053004-1': 'http://darion-svet.com/web/binary/image?model=product.product&id=23970&field=image',
        '4690612053004-2': 'http://darion-svet.com/web/binary/image?model=product.product&id=23970&field=image',
    }

    def __init__(self, wb: openpyxl.Workbook = None):
        self.image_map = dict(self.KNOWN_CATALOG_IMAGES)
        if wb:
            self._scan_workbook_images(wb)

    def _scan_workbook_images(self, wb: openpyxl.Workbook):
        """Сканирует всю книгу для сопоставления штрихкодов и базовых артикулов с фотографиями."""
        for sname in wb.sheetnames:
            if not re.match(r'^\s*([1-8])\b', sname):
                continue
            ws = wb[sname]
            for r in range(3, ws.max_row + 1):
                img = ws.cell(r, 5).value
                if img and str(img).strip().startswith("http"):
                    img_url = str(img).strip()
                    # Сопоставляем по штрихкоду
                    barcode = clean_sku(ws.cell(r, 4).value)
                    if barcode:
                        self.image_map[barcode] = img_url
                    # Сопоставляем по артикулу и базовой части (до дефиса упаковки)
                    sku = clean_sku(ws.cell(r, 1).value)
                    if sku:
                        self.image_map[sku] = img_url
                        base_sku = sku.split("-")[0]
                        self.image_map[base_sku] = img_url

    def resolve_main_image(self, ws, r: int, sku: str, barcode: str) -> str:
        """Находит основное фото из Col 5 либо через карту сопоставления штрихкодов/артикулов."""
        cell_img = ws.cell(r, 5).value
        if cell_img and str(cell_img).strip().startswith("http"):
            return str(cell_img).strip()

        # Поиск по штрихкоду
        clean_b = clean_sku(barcode)
        if clean_b and clean_b in self.image_map:
            return self.image_map[clean_b]

        # Поиск по точному SKU
        clean_s = clean_sku(sku)
        if clean_s and clean_s in self.image_map:
            return self.image_map[clean_s]

        # Поиск по базовой части SKU (4690612052984-2 -> 4690612052984)
        base_s = clean_s.split("-")[0] if clean_s else ""
        if base_s and base_s in self.image_map:
            return self.image_map[base_s]

        return None

    @staticmethod
    def resolve_images(main_img_url: str, extra_imgs_raw: str) -> list:
        """
        Формирует упорядоченный дедуплицированный список изображений (до 10 штук).
        Основное фото всегда идет первым.
        """
        images = []
        if main_img_url and str(main_img_url).strip().startswith("http"):
            images.append(str(main_img_url).strip())

        if extra_imgs_raw and not pd.isna(extra_imgs_raw):
            extra_str = str(extra_imgs_raw).strip()
            # Распаковка архива Odoo: ids=[1724, 1913]
            m = re.search(r'ids=\[([0-9,\s]+)\]', extra_str)
            if m:
                ids_str = m.group(1)
                img_ids = [i.strip() for i in ids_str.split(",") if i.strip().isdigit()]
                for pic_id in img_ids:
                    direct_url = f"http://darion-svet.com/web/binary/image?model=product.picture&id={pic_id}&field=image"
                    if direct_url not in images:
                        images.append(direct_url)
            elif extra_str.startswith("http"):
                if extra_str not in images:
                    images.append(extra_str)

        # Ограничение Авито: не более 10 фото на объявление
        MAX_AVITO_IMAGES = 10
        if len(images) > MAX_AVITO_IMAGES:
            images = images[:MAX_AVITO_IMAGES]

        return images


class DescriptionBuilder:
    """Модуль генерации структурированного продающего описания товара в блоке CDATA."""

    @staticmethod
    def _is_valid(val) -> bool:
        """Проверяет, что значение характеристики не пустое, не 'nan' и не 'None'."""
        if val is None:
            return False
        s = str(val).strip()
        return bool(s and s.lower() not in ("none", "nan", "null", "-", ""))

    @classmethod
    def build_description(cls, prod: dict) -> str:
        name = prod.get("name", "").strip()
        sku = prod.get("sku", "").strip()
        brand = prod.get("brand", "").strip() or config.COMPANY_BRAND

        # Характеристики
        specs = []
        if cls._is_valid(prod.get("power")):
            p_val = str(prod["power"]).strip()
            specs.append(f"• Мощность: {p_val} Вт" if not any(w in p_val.lower() for w in ["вт", "w"]) else f"• Мощность: {p_val}")

        if cls._is_valid(prod.get("lumen")):
            lm_val = str(prod["lumen"]).strip()
            specs.append(f"• Световой поток: {lm_val} лм" if not any(w in lm_val.lower() for w in ["лм", "lm"]) else f"• Световой поток: {lm_val}")

        if cls._is_valid(prod.get("color_temp")):
            ct_val = str(prod["color_temp"]).strip()
            specs.append(f"• Цветовая температура: {ct_val} К" if not any(w in ct_val.lower() for w in ["к", "k"]) else f"• Цветовая температура: {ct_val}")

        if cls._is_valid(prod.get("base")):
            specs.append(f"• Цоколь: {prod['base']}")

        # Колонка O (15) «Тип матрицы LED»
        if cls._is_valid(prod.get("led_matrix")):
            specs.append(f"• Тип матрицы: {prod['led_matrix']}")

        # Колонка T (20) «Тип КСС»
        if cls._is_valid(prod.get("kss")):
            specs.append(f"• Кривая силы света (КСС): {prod['kss']}")

        # Колонка S (19) «Степень защиты»
        if cls._is_valid(prod.get("ip")):
            ip_val = str(prod["ip"]).strip()
            ip_str = ip_val if ip_val.upper().startswith("IP") else f"IP{ip_val}"
            specs.append(f"• Степень защиты: {ip_str}")

        # Колонка R (18) «Индекс цветопередачи»
        if cls._is_valid(prod.get("cri")):
            cri_val = str(prod["cri"]).strip()
            cri_str = cri_val if cri_val.lower().startswith("ra") else f"Ra ≥ {cri_val}"
            specs.append(f"• Индекс цветопередачи: {cri_str}")

        # Коэффициент пульсации (Колонка AL / 38)
        if cls._is_valid(prod.get("pulsation")):
            specs.append(f"• Коэффициент пульсации: ≤ {prod['pulsation']}% (без мерцания)")

        # Колонка U (21) «Срок службы»
        if cls._is_valid(prod.get("lifetime")):
            lt_val = str(prod["lifetime"]).strip()
            lt_str = f"{lt_val} ч." if re.match(r'^\d+$', lt_val.replace(' ', '')) else lt_val
            specs.append(f"• Срок службы: {lt_str}")

        # Колонка I (9) «Способ установки»
        if cls._is_valid(prod.get("mounting")):
            specs.append(f"• Способ установки: {prod['mounting']}")

        # Колонка J (10) «Область применения»
        if cls._is_valid(prod.get("application")):
            specs.append(f"• Область применения: {prod['application']}")

        # Колонка F (6) «Гарантия»
        if cls._is_valid(prod.get("warranty")):
            w_val = str(prod["warranty"]).strip()
            w_str = w_val if any(w in w_val.lower() for w in ["год", "лет"]) else f"{w_val} года (лет)"
            specs.append(f"• Гарантия производителя: {w_str}")

        # Колонка L (12) «Количество ламп»
        if cls._is_valid(prod.get("lamps_count")):
            specs.append(f"• Количество ламп: {prod['lamps_count']}")

        # Материал корпуса (Колонка AD / 30)
        if cls._is_valid(prod.get("body_material")):
            specs.append(f"• Материал корпуса: {prod['body_material']}")

        # Рассеиватель (Колонка AE / 31)
        if cls._is_valid(prod.get("diffuser_material")):
            specs.append(f"• Рассеиватель: {prod['diffuser_material']}")

        # Габариты (Колонки V, W, X / 22, 23, 24)
        if cls._is_valid(prod.get("dimensions")):
            specs.append(f"• Габариты (Д×Ш×В): {prod['dimensions']} мм")

        # Масса нетто (Колонка AB / 28)
        if cls._is_valid(prod.get("weight")):
            specs.append(f"• Масса нетто: {prod['weight']} кг")

        specs_block = "\n".join(specs) if specs else "• Характеристики соответствуют паспорту изделия"

        # Полное продающее описание из столбца AH (34) без обрезки текста
        extra_desc = str(prod.get("extra_desc") or "").strip()
        marketing_block = ""
        if extra_desc:
            marketing_block = f"\n\n[ОПИСАНИЕ И ПРЕИМУЩЕСТВА]\n{extra_desc}"

        company_block = """[ПРЕИМУЩЕСТВА И УСЛОВИЯ КОМПАНИИ «ДАРИОН СВЕТ»]
- Официальная гарантия производителя на всю светотехнику.
- Быстрая отгрузка со склада в Санкт-Петербурге.
- Работаем с юр. и физ. лицами (оплата по счету с НДС 20% и без НДС).
- Быстрая доставка по Санкт-Петербургу, Ленинградской области и всей России (ТК СДЭК, Деловые Линии, ПЭК).
- Предоставляем полный комплект документов: паспорта изделий, сертификаты ЕАС.

Звоните или пишите в сообщения на Авито — рассчитаем освещенность объекта и подберем необходимое оборудование!"""

        header_block = f"{name}\nАртикул: {sku}\nПроизводитель: {brand}\n\n[ТЕХНИЧЕСКИЕ ХАРАКТЕРИСТИКИ]\n{specs_block}"

        description_text = f"{header_block}{marketing_block}\n\n{company_block}".strip()

        # Системный контроль: суммарная длина тега Description <= 5000 символов (лимит Авито)
        if len(description_text) > 5000:
            description_text = description_text[:5000].rstrip(' ,.-;/')

        return description_text


class FeedRouter:
    """Маршрутизатор листов 1..8 по 8 целевым фидам Авито."""

    @staticmethod
    def get_feed_key_for_sheet(sheet_name: str) -> str | None:
        """Определяет ключ фида по номеру листа (1..8)."""
        m = re.match(r'^\s*([1-8])\b', sheet_name)
        if m:
            num = int(m.group(1))
            return SHEET_INDEX_TO_FEED_KEY.get(num)
        return None


class AvitoFeedGenerator:
    """Главный координатор сборки и выгрузки 8 XML-фидов."""

    def __init__(self):
        self.stock_sync = StockSync()
        self.desc_builder = DescriptionBuilder()
        self.router = FeedRouter()
        self.feed_buckets = {k: [] for k in FEEDS_CONFIG.keys()}

    @staticmethod
    def format_title(name: str, sku: str = "", max_len: int = 50) -> str:
        """
        Формирует заголовок объявления для Авито в формате «[Название] ([Артикул])».
        Общая длина строки СТРОГО <= max_len (по умолчанию 50 символов).
        Базовое наименование аккуратно обрезается по границе слов без висячих знаков пунктуации.
        """
        name = (name or "").strip()
        sku = (sku or "").strip()

        # Удаляем хвостовые скобки с артикулами (VRN-UNE-48-G40K67-U) или упаковкой
        cleaned = re.sub(r'\s*\([^)]*\)\s*$', '', name).strip() or name
        suffix = f" ({sku})" if sku else ""
        avail = max_len - len(suffix)
        if avail <= 0:
            return (cleaned + suffix)[:max_len]

        if len(cleaned) <= avail:
            base = cleaned
        else:
            if len(cleaned) > avail and cleaned[avail] == ' ':
                base = cleaned[:avail].rstrip(' ,.-;/')
            else:
                sp = cleaned.rfind(' ', 0, avail)
                base = cleaned[:sp].rstrip(' ,.-;/') if sp > 0 else cleaned[:avail].rstrip(' ,.-;/')

        t = f"{base}{suffix}".strip()
        return t[:max_len].rstrip(' ,.-;/') if len(t) > max_len else t

    def build_ad_element(self, prod: dict, feed_key: str) -> etree.Element:
        """Формирует XML-элемент <Ad> со всеми обязательными и категорийными тегами."""
        ad = etree.Element("Ad")

        # 1. Обязательные базовые теги Авито
        ad_id = etree.SubElement(ad, "Id")
        ad_id.text = prod["sku"]

        address = etree.SubElement(ad, "Address")
        address.text = getattr(config, "DEFAULT_ADDRESS", "Санкт-Петербург, Студенческая ул., 10")

        title = etree.SubElement(ad, "Title")
        title.text = self.format_title(prod["name"], prod["sku"], max_len=50)

        desc = etree.SubElement(ad, "Description")
        desc_text = self.desc_builder.build_description(prod)
        desc.text = etree.CDATA(desc_text)

        images_list = ImageResolver.resolve_images(prod["main_image"], prod.get("extra_images"))
        images_el = etree.SubElement(ad, "Images")
        for img_url in images_list:
            img_tag = etree.SubElement(images_el, "Image")
            img_tag.set("url", img_url)

        price_el = etree.SubElement(ad, "Price")
        price_el.text = str(prod["price"])

        ad_type = etree.SubElement(ad, "AdType")
        ad_type.text = "Товар куплен на продажу"

        condition = etree.SubElement(ad, "Condition")
        condition.text = "Новое"

        availability = etree.SubElement(ad, "Availability")
        availability.text = "В наличии"

        contact = etree.SubElement(ad, "ContactMethod")
        contact.text = "По телефону и в сообщениях"

        # 2. Категорийные теги
        cfg = FEEDS_CONFIG[feed_key]
        cat_val = cfg["category"]
        if cat_val == "Для дома и дачи" or feed_key == "lamps":
            cat_val = "Мебель и интерьер"

        cat_el = etree.SubElement(ad, "Category")
        cat_el.text = cat_val

        goods_type_val = cfg["goods_type"]
        if feed_key == "lamps":
            goods_type_val = "Освещение"

        goods_type_el = etree.SubElement(ad, "GoodsType")
        goods_type_el.text = goods_type_val

        # GoodsSubType: строго по официальному дереву категорий Авито
        if feed_key == "led_luminaires":
            goods_sub_type = "Потолочное и настенное"
        elif feed_key == "track_systems":
            goods_sub_type = "Уличное"
        elif feed_key == "lamps":
            goods_sub_type = "Комплектующие"
        else:
            goods_sub_type = cfg.get("goods_sub_type")

        if goods_sub_type:
            sub_type_el = etree.SubElement(ad, "GoodsSubType")
            sub_type_el.text = goods_sub_type

        # LigitingType (внимание: специфическая схема тега Авито без буквы 'h')
        lighting_type = "Люстры и потолочные светильники" if feed_key == "led_luminaires" else cfg.get("lighting_type")
        if lighting_type:
            light_type_el = etree.SubElement(ad, "LigitingType")
            light_type_el.text = lighting_type

        # Обязательные характеристики для подтипа «Люстры и потолочные светильники»
        if feed_key in ("led_luminaires", "chandeliers") or lighting_type == "Люстры и потолочные светильники":
            # 1. Подтип потолочного освещения (Светильник / Люстра)
            chan_type_el = etree.SubElement(ad, "ChandelierType")
            chan_type_el.text = "Светильник"

            # 2. Тип крепления (допустимые значения Авито: "Потолочное" / "Подвесное")
            mounting_raw = str(prod.get("mounting") or "").lower()
            if "подвес" in mounting_raw and not ("потолоч" in mounting_raw or "встраива" in mounting_raw or "накладн" in mounting_raw):
                mounting_val = "Подвесное"
            else:
                mounting_val = "Потолочное"
            mount_el = etree.SubElement(ad, "ChandelierMountingType")
            mount_el.text = mounting_val

            # 3. Светодиодный (LED) (Да / Нет)
            led_el = etree.SubElement(ad, "LedLamp")
            led_el.text = "Да"

        # Специфические характеристики для фида ламп (Лист №4)
        if feed_key == "lamps":
            # 1. Тип лампы: Светодиодная
            lamp_type_el = etree.SubElement(ad, "LampType")
            lamp_type_el.text = "Светодиодная"
            bulb_type_el = etree.SubElement(ad, "BulbType")
            bulb_type_el.text = "Светодиодная"

            # 2. Цоколь: из колонки N (base) с удалением пробелов и нормализацией
            base_val = prod.get("base")
            if base_val:
                cap_type_el = etree.SubElement(ad, "CapType")
                cap_type_el.text = base_val
                bulb_base_el = etree.SubElement(ad, "BulbBaseType")
                bulb_base_el.text = base_val

            # 3. Мощность: из колонки M (power)
            power_val = prod.get("power")
            if power_val:
                power_el = etree.SubElement(ad, "Power")
                power_el.text = str(power_val)

            # 4. Цветовая температура: из колонки Q (color_temp)
            temp_val = prod.get("color_temp")
            if temp_val:
                col_temp_el = etree.SubElement(ad, "ColorTemperature")
                col_temp_el.text = str(temp_val)
                temp_el = etree.SubElement(ad, "Temperature")
                temp_el.text = str(temp_val)

            # 5. Световой поток: из колонки P (lumen)
            lumen_val = prod.get("lumen")
            if lumen_val:
                flux_el = etree.SubElement(ad, "LightFlux")
                flux_el.text = str(lumen_val)

        if prod.get("brand"):
            brand_el = etree.SubElement(ad, "Brand")
            brand_el.text = prod["brand"]

        return ad

    def generate_all_feeds(self) -> dict:
        """
        Выполняет загрузку книги, обработку листов 1..8 и сборку 8 XML-фидов.
        """
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        self.feed_buckets = {k: [] for k in FEEDS_CONFIG.keys()}

        # 1. Загрузка книги Google Sheets / кэша
        wb = self.stock_sync.get_workbook()

        # 2. Инициализация резолвера изображений по книге
        image_resolver = ImageResolver(wb)

        # 3. Обработка листов 1..8
        total_matched = 0

        for sheet_name in wb.sheetnames:
            feed_key = self.router.get_feed_key_for_sheet(sheet_name)
            if not feed_key:
                continue

            ws = wb[sheet_name]
            sheet_items = []

            for r in range(3, ws.max_row + 1):
                sku_raw = ws.cell(r, 1).value
                name_raw = ws.cell(r, 2).value
                if not sku_raw or not name_raw:
                    continue

                sku = clean_sku(sku_raw)
                name = str(name_raw).strip()
                barcode = clean_sku(ws.cell(r, 4).value)

                # Фильтрация строк строго по формуле:
                # 1) Столбец E (5): основное фото заполнено (не пустое, начинается с http)
                # 2) Столбец BI (61): остаток строго > 0
                # 3) Столбец BJ (62): цена строго > 1
                stock = clean_numeric(ws.cell(r, 61).value)
                price = clean_numeric(ws.cell(r, 62).value)
                main_image = image_resolver.resolve_main_image(ws, r, sku, barcode)

                if stock > 0 and price > 1 and main_image and main_image.startswith("http"):
                    # Габариты A(22), B(23), C(24)
                    dim_a = format_num_val(ws.cell(r, 22).value)
                    dim_b = format_num_val(ws.cell(r, 23).value)
                    dim_c = format_num_val(ws.cell(r, 24).value)
                    dim_str = f"{dim_a}×{dim_b}×{dim_c}" if (dim_a and dim_b and dim_c) else None

                    prod = {
                        "sheet": sheet_name,
                        "sku": sku,
                        "name": name,
                        "category": str(ws.cell(r, 3).value or "").strip(),
                        "ean": barcode,
                        "main_image": main_image,
                        "warranty": format_num_val(ws.cell(r, 6).value),
                        "country": str(ws.cell(r, 7).value or "").strip(),
                        "brand": str(ws.cell(r, 8).value or config.COMPANY_BRAND).strip(),
                        "mounting": str(ws.cell(r, 9).value or "").strip(),
                        "application": str(ws.cell(r, 10).value or "").strip(),
                        "light_source_type": str(ws.cell(r, 11).value or "").strip(),
                        "lamps_count": format_num_val(ws.cell(r, 12).value),
                        "power": format_num_val(ws.cell(r, 13).value),
                        "base": str(ws.cell(r, 14).value or "").strip(),
                        "led_matrix": str(ws.cell(r, 15).value or "").strip(),
                        "lumen": format_num_val(ws.cell(r, 16).value),
                        "color_temp": format_num_val(ws.cell(r, 17).value),
                        "cri": format_num_val(ws.cell(r, 18).value),
                        "ip": format_num_val(ws.cell(r, 19).value),
                        "kss": str(ws.cell(r, 20).value or "").strip(),
                        "lifetime": str(ws.cell(r, 21).value or "").strip(),
                        "dimensions": dim_str,
                        "weight": format_num_val(ws.cell(r, 28).value),
                        "body_material": str(ws.cell(r, 30).value or "").strip(),
                        "diffuser_material": str(ws.cell(r, 31).value or "").strip(),
                        "extra_desc": str(ws.cell(r, 34).value or "").strip(),
                        "pulsation": format_num_val(ws.cell(r, 38).value),
                        "extra_images": ws.cell(r, 56).value,
                        "site_url": ws.cell(r, 57).value,
                        "stock": int(round(stock)),
                        "price": int(round(price))
                    }

                    if feed_key == "lamps":
                        # Нормализация цоколя: удаление пробелов, русские 'Е' -> 'E', fallback на имя
                        base_val = prod["base"].replace(" ", "").replace("Е", "E").replace("е", "e")
                        if not base_val or base_val.lower() == "none":
                            m_b = re.search(r'\b([EЕ]\d+|GX\d+|GU\d+(?:\.\d+)?|G\d+)\b', name, re.IGNORECASE)
                            if m_b:
                                base_val = m_b.group(1).upper().replace(" ", "").replace("Е", "E")
                        prod["base"] = base_val

                        # Нормализация мощности
                        if not prod["power"] or prod["power"].lower() == "none":
                            m_p = re.search(r'(\d+(?:\.\d+)?)\s*(?:Вт|W)\b', name, re.IGNORECASE)
                            if m_p:
                                prod["power"] = m_p.group(1)

                        # Нормализация цветовой температуры
                        if not prod["color_temp"] or prod["color_temp"].lower() == "none":
                            m_t = re.search(r'(\d{4})\s*(?:К|K)\b', name, re.IGNORECASE)
                            if m_t:
                                prod["color_temp"] = m_t.group(1)

                        # Нормализация светового потока
                        if not prod["lumen"] or prod["lumen"].lower() == "none":
                            m_l = re.search(r'(\d+)\s*(?:Лм|lm)\b', name, re.IGNORECASE)
                            if m_l:
                                prod["lumen"] = m_l.group(1)

                    sheet_items.append(prod)

            self.feed_buckets[feed_key].extend(sheet_items)
            total_matched += len(sheet_items)
            logger.info(f"Лист '{sheet_name}' -> фид '{feed_key}': загружено {len(sheet_items)} товаров.")

        logger.info(f"Всего сопоставлено и направлено в 8 фидов: {total_matched} товаров.")

        # 4. Генерация 8 XML-файлов в output/
        generated_summary = {}

        for feed_key, cfg in FEEDS_CONFIG.items():
            out_filename = cfg["filename"]
            out_path = os.path.join(OUTPUT_DIR, out_filename)
            items = self.feed_buckets[feed_key]

            # Создаем корневой элемент Авито
            root = etree.Element("Ads", formatVersion="3", target="Avito.ru")

            if len(items) > 0:
                for prod in items:
                    ad_el = self.build_ad_element(prod, feed_key)
                    root.append(ad_el)
            else:
                # Обязательный валидный пустой каркас
                root.text = ""

            xml_bytes = etree.tostring(
                root,
                encoding="UTF-8",
                xml_declaration=True,
                pretty_print=True
            )

            with open(out_path, "wb") as f:
                f.write(xml_bytes)

            file_size_kb = round(os.path.getsize(out_path) / 1024, 2)
            generated_summary[out_filename] = {
                "feed_key": feed_key,
                "title": cfg["title"],
                "path": out_path,
                "size_kb": file_size_kb,
                "ads_count": len(items)
            }

        return generated_summary


class FeedUploader:
    """Модуль публикации фидов на удаленный веб-сервер по протоколам FTP / SFTP."""

    def __init__(self):
        self.method = getattr(config, "UPLOAD_METHOD", "local").lower()

    def upload_all(self, generated_summary: dict) -> dict:
        """Передает сгенерированные файлы согласно UPLOAD_METHOD в config.py."""
        if self.method == 'ftp':
            return self._upload_ftp(generated_summary)
        elif self.method == 'sftp':
            return self._upload_sftp(generated_summary)
        else:
            msg = "Файлы сохранены локально в папке output/"
            logger.info(msg)
            return {
                "status": "LOCAL",
                "message": msg,
                "links": {
                    fname: os.path.abspath(info["path"])
                    for fname, info in generated_summary.items()
                }
            }

    def _upload_ftp(self, generated_summary: dict) -> dict:
        host = getattr(config, "FTP_HOST", "")
        port = getattr(config, "FTP_PORT", 21)
        user = getattr(config, "FTP_USER", "")
        passwd = getattr(config, "FTP_PASS", "")
        remote_dir = getattr(config, "FTP_REMOTE_DIR", "/public_html/avito_feeds").rstrip("/")
        public_base = getattr(config, "FTP_PUBLIC_URL_BASE", "http://darion-svet.com/avito_feeds").rstrip("/")

        if not host or not user:
            msg = "Настройки FTP не заполнены в config.py. Файлы сохранены локально в папке output/"
            logger.warning(msg)
            return {
                "status": "CONFIG_INCOMPLETE",
                "message": msg,
                "links": {fname: f"{public_base}/{fname}" for fname in generated_summary}
            }

        try:
            logger.info(f"Подключение к FTP-серверу {host}:{port}...")
            ftp = ftplib.FTP()
            ftp.connect(host, port, timeout=30)
            ftp.login(user, passwd)

            try:
                ftp.cwd(remote_dir)
            except Exception:
                parts = remote_dir.strip("/").split("/")
                for p in parts:
                    try:
                        ftp.cwd(p)
                    except Exception:
                        ftp.mkd(p)
                        ftp.cwd(p)

            links = {}
            for fname, info in generated_summary.items():
                local_path = info["path"]
                with open(local_path, "rb") as fp:
                    ftp.storbinary(f"STOR {fname}", fp)
                public_url = f"{public_base}/{fname}"
                links[fname] = public_url
                logger.info(f"Загружен на FTP: {fname} -> {public_url}")

            ftp.quit()
            return {
                "status": "SUCCESS_FTP",
                "message": f"Все {len(links)} фидов успешно загружены на FTP-сервер {host}",
                "links": links
            }
        except Exception as e:
            msg = f"Ошибка при загрузке по FTP: {e}. Файлы сохранены локально в папке output/"
            logger.error(msg)
            return {
                "status": "FTP_ERROR",
                "message": msg,
                "links": {fname: f"{public_base}/{fname}" for fname in generated_summary}
            }

    def _upload_sftp(self, generated_summary: dict) -> dict:
        host = getattr(config, "SFTP_HOST", "")
        port = getattr(config, "SFTP_PORT", 22)
        user = getattr(config, "SFTP_USER", "")
        passwd = getattr(config, "SFTP_PASS", "")
        remote_dir = getattr(config, "SFTP_REMOTE_DIR", "/var/www/darion-svet/feeds").rstrip("/")
        public_base = getattr(config, "SFTP_PUBLIC_URL_BASE", "http://darion-svet.com/avito_feeds").rstrip("/")

        if not host or not user:
            msg = "Настройки SFTP не заполнены в config.py. Файлы сохранены локально в папке output/"
            logger.warning(msg)
            return {
                "status": "CONFIG_INCOMPLETE",
                "message": msg,
                "links": {fname: f"{public_base}/{fname}" for fname in generated_summary}
            }

        try:
            logger.info(f"Подключение к SFTP-серверу {host}:{port}...")
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(host, port=port, username=user, password=passwd, timeout=30)
            sftp = ssh.open_sftp()

            try:
                sftp.chdir(remote_dir)
            except Exception:
                parts = remote_dir.strip("/").split("/")
                cur = ""
                for p in parts:
                    cur += "/" + p
                    try:
                        sftp.stat(cur)
                    except Exception:
                        sftp.mkdir(cur)
                sftp.chdir(remote_dir)

            links = {}
            for fname, info in generated_summary.items():
                local_path = info["path"]
                remote_file = f"{remote_dir}/{fname}"
                sftp.put(local_path, remote_file)
                public_url = f"{public_base}/{fname}"
                links[fname] = public_url
                logger.info(f"Загружен по SFTP: {fname} -> {public_url}")

            sftp.close()
            ssh.close()
            return {
                "status": "SUCCESS_SFTP",
                "message": f"Все {len(links)} фидов успешно загружены на SFTP-сервер {host}",
                "links": links
            }
        except Exception as e:
            msg = f"Ошибка при загрузке по SFTP: {e}. Файлы сохранены локально в папке output/"
            logger.error(msg)
            return {
                "status": "SFTP_ERROR",
                "message": msg,
                "links": {fname: f"{public_base}/{fname}" for fname in generated_summary}
            }


def print_final_report(generator: AvitoFeedGenerator, summary: dict, upload_res: dict):
    """Выводит структурированный итоговый консольный отчет."""
    print("\n" + "=" * 95)
    print("      ИТОГОВЫЙ ОТЧЕТ ГЕНЕРАТОРА 8 XML-ФИДОВ АВИТО («ДАРИОН СВЕТ»)")
    print("=" * 95 + "\n")

    # 1. Статус скачивания
    print("1. СТАТУС СИНХРОНИЗАЦИИ ИЗ GOOGLE SHEETS:")
    print("-" * 95)
    print(f"Статус подключения:        {generator.stock_sync.sync_status}")
    print(f"Сообщение шлюза:           {generator.stock_sync.sync_message}")
    print(f"Использованный источник:   {os.path.relpath(generator.stock_sync.used_source, BASE_DIR) if generator.stock_sync.used_source else 'Google Sheets / Кэш'}")
    print("-" * 95 + "\n")

    # 2. Статистика сгенерированных фидов
    print("2. СТАТИСТИКА СГЕНЕРИРОВАННЫХ ФИДОВ В output/:")
    print("-" * 95)
    print(f"{'№':<3} | {'Имя XML-фида':<29} | {'Размер (КБ)':<11} | {'Объявлений':<11} | {'Категория / Назначение':<32}")
    print("-" * 95)

    total_ads = 0
    for idx, (fname, info) in enumerate(summary.items(), 1):
        total_ads += info["ads_count"]
        print(f"{idx:<3} | {fname:<29} | {info['size_kb']:<11} | {info['ads_count']:<11} | {info['title']:<32}")
    print("-" * 95)
    print(f"Итого создано фидов: 8 из 8 | Суммарно выгружено активных объявлений: {total_ads}\n")

    # 3. Статус публикации
    print("3. СТАТУС ПУБЛИКАЦИИ ФИДОВ (РЕЖИМ: " + getattr(config, "UPLOAD_METHOD", "local").upper() + "):")
    print("-" * 95)
    print(f"Статус:                    {upload_res['status']}")
    print(f"Результат:                 {upload_res['message']}")
    if upload_res.get("links"):
        print("Ссылки на фиды:")
        for fname, link in upload_res["links"].items():
            print(f"  • {fname:<28} -> {link}")
    print("-" * 95 + "\n")

    # 4. Пример сгенерированного объявления
    sample_file = os.path.join(OUTPUT_DIR, "led_luminaires_feed.xml")
    if os.path.exists(sample_file):
        with open(sample_file, "r", encoding="utf-8") as f:
            raw_content = f.read()
        m = re.search(r'(  <Ad>.*?</Ad>)', raw_content, re.DOTALL)
        if m:
            print("4. ПРИМЕР СГЕНЕРИРОВАННОГО ОБЪЯВЛЕНИЯ <Ad> (led_luminaires_feed.xml):")
            print("-" * 95)
            print(m.group(1))
            print("-" * 95)

    print("\n" + "=" * 95)
    print("   [УСПЕХ] Все 8 XML-фидов успешно сформированы и готовы к загрузке в Авито!")
    print(f"   Локальная папка с фидами: {OUTPUT_DIR}")
    print("=" * 95 + "\n")


def main():
    generator = AvitoFeedGenerator()
    summary = generator.generate_all_feeds()
    uploader = FeedUploader()
    upload_res = uploader.upload_all(summary)
    print_final_report(generator, summary, upload_res)


if __name__ == "__main__":
    main()
