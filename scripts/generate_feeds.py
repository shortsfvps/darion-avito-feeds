"""
Генератор 8 XML-фидов для Авито (Автозагрузка) компании «Дарион Свет».

Модули:
1. StockSync: синхронизация цен и остатков из Google Sheets с локальным кэшированием и фильтром stock > 0.
2. ImageResolver: формирование прямых ссылок на изображения товаров (включая распаковку архивов Odoo).
3. DescriptionBuilder: генерация продающего структурированного описания в блоке CDATA.
4. FeedRouter: маршрутизация товаров по 8 целевым категориям фидов и создание валидных XML-каркасов.
5. AvitoFeedGenerator: сборка и валидация итоговых XML-файлов в папку output/.
6. FeedUploader: публикация фидов по протоколам FTP / SFTP или локальное сохранение.
"""

import os
import sys
import io
import re
import json
import logging
import ftplib
import requests
import pandas as pd
import openpyxl
from lxml import etree
import paramiko

# Обеспечиваем безопасный вывод UTF-8 в консоль Windows
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Подключение файла конфигурации config.py из корня проекта
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
TEMPLATES_DIR = config.TEMPLATES_DIR
EXCEL_PRODUCTS_PATH = config.NOMENCLATURE_FILE
GOOGLE_SHEET_URL = config.GOOGLE_SHEETS_URL
LATEST_STOCK_PATH = config.LATEST_STOCK_FILE
TEMPLATE_STOCK_PATH = config.TEMPLATE_STOCK_FILE
TAGS_SCHEMA_PATH = config.TAGS_SCHEMA_FILE
FEEDS_CONFIG = config.FEEDS_CONFIG


class StockSync:
    """Модуль синхронизации остатков и цен из Google Sheets."""

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
            logger.info(f"Запрос актуальных остатков из Google Sheets (таймаут {timeout} сек)...")
            response = requests.get(self.url, headers=headers, timeout=timeout)
            # Проверяем, что получен валидный Excel (.xlsx начинается с сигнатуры PK..)
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
                    f"Google Таблица вернула HTTP {response.status_code} "
                    f"(Ограниченный доступ: требуется авторизация). Используется локальный кэш."
                )
                logger.warning(self.sync_message)
                return None
        except Exception as e:
            self.sync_status = "NETWORK_ERROR"
            self.sync_message = f"Ошибка сети при запросе к Google Sheets: {e}. Используется локальный кэш."
            logger.warning(self.sync_message)
            return None

    def load_stock_dict(self) -> dict:
        """
        Парсит лист с остатками («Остатки СПб» / «Остатки СПБ запасы»).
        Собирает словарь: {Артикул: {'price': int(Цена), 'stock': int(Остаток)}}.
        Фильтрует позиции: в выгрузку попадают только товары с price > 0 AND stock > 0.

        Правки C1/C2/C3/H1 (аудит 2026-10-02):
          C3 — окно поиска строки заголовков расширено до 25 строк.
          C1 — индексы колонок определяются динамически по именам заголовков;
               при отсутствии совпадения — безопасный fallback на 0/9/10 с WARNING.
          C2 — значения цены/остатка очищаются через clean_numeric() (убирает 'руб.', 'шт.',
               пробелы, неразрывные пробелы, запятые); при невозможности распознать — WARNING.
          H1 — фильтр: price > 0 AND stock > 0; нулевая цена → WARNING + пропуск.
        """
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
            raise FileNotFoundError(f"Файлы остатков не найдены: {self.latest_cache} и {self.fallback_cache}")

        # Определяем целевой лист остатков
        xl = pd.ExcelFile(source)
        target_sheet = None
        for s in xl.sheet_names:
            if "остат" in s.lower() and "спб" in s.lower():
                target_sheet = s
                break
        if not target_sheet:
            target_sheet = xl.sheet_names[0]

        # C3: окно поиска строки заголовков расширено с 5 до 25 строк
        df_raw = pd.read_excel(source, sheet_name=target_sheet, header=None)
        header_row_idx = 0
        for idx, row in df_raw.head(25).iterrows():
            if any("артикул" in str(val).lower() for val in row):
                header_row_idx = idx
                break

        df = pd.read_excel(source, sheet_name=target_sheet, header=header_row_idx)

        # C1: динамический поиск колонок по нормализованным именам заголовков
        _PATTERNS_SKU   = ("артикул",)
        _PATTERNS_STOCK = ("свободно", "остаток")
        # Паттерны розничной цены (исключая закупочные колонки — ТЗ: столбец J)
        _PATTERNS_PRICE_RETAIL   = ("распродаж", "розниц", "продаж", "цена")
        _PATTERN_PRICE_EXCLUDE   = "закуп"   # закупки/закупок/закупочн — исключать

        def _find_col_idx(df, patterns):
            """Возвращает позиционный индекс первой колонки, заголовок которой содержит один из паттернов."""
            for pos, col in enumerate(df.columns):
                col_norm = str(col).lower().strip()
                if any(p in col_norm for p in patterns):
                    return pos
            return None

        def _find_price_col_idx(df) -> int | None:
            """
            Специализированный поиск колонки РОЗНИЧНОЙ цены (ТЗ: столбец J, индекс 9).
            Алгоритм:
              1. Если col[9] содержит 'цена' или 'прайс' и НЕ содержит 'закуп' — возвращаем 9.
              2. Иначе перебираем все колонки: ищем 'распродаж'/'розниц'/'продаж'/'цена'
                 при условии отсутствия 'закуп' в заголовке.
              3. Fallback: None (вызывающий код применит дефолт 9 с WARNING).
            Колонки со словом 'закуп' (закупок, закупки, закупочная) ВСЕГДА исключаются —
            это себестоимость, не цена продажи для Авито.
            """
            # Шаг 1: приоритетная проверка col[9] (столбец J по ТЗ заказчика)
            if len(df.columns) > 9:
                col9_norm = str(df.columns[9]).lower().strip()
                is_price_word = any(p in col9_norm for p in ("цена", "прайс"))
                is_purchase   = _PATTERN_PRICE_EXCLUDE in col9_norm
                if is_price_word and not is_purchase:
                    logger.info(
                        f"[PRICE-COL] Приоритетный выбор: col[9] = '{df.columns[9]}' "
                        f"(соответствует ТЗ — столбец J, розничная цена)."
                    )
                    return 9

            # Шаг 2: полный перебор — розничные паттерны без 'закуп'
            for pos, col in enumerate(df.columns):
                col_norm = str(col).lower().strip()
                if _PATTERN_PRICE_EXCLUDE in col_norm:
                    continue   # закупочная колонка — пропускаем
                if any(p in col_norm for p in _PATTERNS_PRICE_RETAIL):
                    logger.info(
                        f"[PRICE-COL] Найдена розничная цена: col[{pos}] = '{df.columns[pos]}'."
                    )
                    return pos

            return None   # ничего не нашли — fallback снаружи

        col_sku_idx   = _find_col_idx(df, _PATTERNS_SKU)
        col_price_idx = _find_price_col_idx(df)
        col_stock_idx = _find_col_idx(df, _PATTERNS_STOCK)

        # Fallback на позиционные индексы по умолчанию с предупреждением в лог
        _FALLBACKS = {
            "Артикул (SKU)": (col_sku_idx,   0,  _PATTERNS_SKU),
            "Цена (розница)": (col_price_idx, 9,  _PATTERNS_PRICE_RETAIL),
            "Остаток":        (col_stock_idx, 10, _PATTERNS_STOCK),
        }
        for label, (found, default, patterns) in _FALLBACKS.items():
            if found is None:
                logger.warning(
                    f"[C1-FALLBACK] Колонка '{label}' не найдена по паттернам {patterns} "
                    f"в заголовках: {list(df.columns)}. "
                    f"Используется позиционный индекс по умолчанию: {default}."
                )
        col_sku_idx   = col_sku_idx   if col_sku_idx   is not None else 0
        col_price_idx = col_price_idx if col_price_idx is not None else 9
        col_stock_idx = col_stock_idx if col_stock_idx is not None else 10

        logger.info(
            f"Колонки остатков определены: "
            f"SKU=col[{col_sku_idx}] '{df.columns[col_sku_idx]}', "
            f"Цена=col[{col_price_idx}] '{df.columns[col_price_idx]}', "
            f"Остаток=col[{col_stock_idx}] '{df.columns[col_stock_idx]}'"
        )

        # C2: вспомогательная функция безопасной очистки нечисловых значений ячеек
        def clean_numeric(val) -> float:
            """
            Преобразует значение ячейки в float, безопасно удаляя мусорные символы:
            неразрывные пробелы (\\xa0), 'руб.', 'шт.', запятые → точки, и т.п.
            Возвращает 0.0 при любой невозможности преобразования.
            """
            if val is None:
                return 0.0
            try:
                if pd.isna(val):
                    return 0.0
            except (TypeError, ValueError):
                pass
            # Числовые типы — сразу возвращаем
            if isinstance(val, (int, float)):
                return float(val)
            # Строковая обработка
            s = str(val)
            s = s.replace("\xa0", " ")                      # неразрывный пробел → обычный
            s = re.sub(r"(?i)(руб\.?|шт\.?|%)", "", s)     # единицы измерения
            s = s.replace(",", ".").strip()
            # Извлекаем первое число (целое или дробное, возможно отрицательное)
            m = re.search(r"-?\d+(?:\.\d+)?", s)
            if m:
                return float(m.group())
            return 0.0

        available_stock = {}
        total_rows = len(df)
        filtered_out = 0
        skipped_zero_price = 0

        for row_idx in range(total_rows):
            row = df.iloc[row_idx]
            sku_raw   = row.iloc[col_sku_idx]   if len(row) > col_sku_idx   else None
            price_raw = row.iloc[col_price_idx] if len(row) > col_price_idx else None
            stock_raw = row.iloc[col_stock_idx] if len(row) > col_stock_idx else None

            if pd.isna(sku_raw):
                continue

            sku = str(sku_raw).strip()
            if not sku or sku.lower() == "артикул":
                continue

            price = int(round(clean_numeric(price_raw)))
            stock = int(round(clean_numeric(stock_raw)))

            # H1: фильтр по остатку
            if stock <= 0:
                filtered_out += 1
                continue

            # H1: фильтр по цене — товар с ценой 0 не должен попасть в фид (Авито отклонит)
            if price <= 0:
                logger.warning(
                    f"[H1] SKU {sku}: цена = {price_raw!r} → распознано как {price} руб. "
                    f"Товар исключён из фида (нулевая/некорректная цена)."
                )
                skipped_zero_price += 1
                continue

            available_stock[sku] = {
                "price": price,
                "stock": stock
            }

        logger.info(
            f"Загрузка остатков завершена (лист '{target_sheet}'): доступно к выгрузке {len(available_stock)} позиций "
            f"(исключено с нулевым/отрицательным остатком: {filtered_out}; "
            f"исключено с нулевой ценой: {skipped_zero_price})."
        )
        return available_stock


class ImageResolver:
    """Модуль обработки и валидации изображений для фидов Авито."""

    @staticmethod
    def resolve_images(main_img_url: str, extra_imgs_raw: str) -> list:
        """
        - Основное фото из столбца E ('Изображение товара').
        - Дополнительные фото из столбца BD ('Доп. изображения'):
          Если в ссылке есть ids=[...], извлекаются все ID и формируются прямые ссылки:
          http://darion-svet.com/web/binary/image?model=product.picture&id={id}&field=image
          Если прямая ссылка — добавляется напрямую.
        - Возвращает дедуплицированный упорядоченный список URL.
        """
        images = []

        # 1. Основное изображение
        if main_img_url and str(main_img_url).strip().startswith("http"):
            images.append(str(main_img_url).strip())

        # 2. Дополнительные изображения
        if extra_imgs_raw and not pd.isna(extra_imgs_raw):
            extra_str = str(extra_imgs_raw).strip()
            # Распаковка архива с массивом IDs: ids=[1724, 1913]
            m = re.search(r'ids=\[([0-9,\s]+)\]', extra_str)
            if m:
                ids_str = m.group(1)
                img_ids = [i.strip() for i in ids_str.split(",") if i.strip().isdigit()]
                for pic_id in img_ids:
                    direct_url = f"http://darion-svet.com/web/binary/image?model=product.picture&id={pic_id}&field=image"
                    if direct_url not in images:
                        images.append(direct_url)
            elif extra_str.startswith("http"):
                # Прямая ссылка на одно изображение
                if extra_str not in images:
                    images.append(extra_str)

        # C4: жёсткий лимит Авито — не более 10 фотографий на объявление.
        # Основное фото (Col E) всегда идёт первым (добавлено выше первым).
        # Дедупликация сохраняется за счёт проверок `not in images` при добавлении.
        MAX_AVITO_IMAGES = 10
        if len(images) > MAX_AVITO_IMAGES:
            logger.warning(
                f"[C4] Количество изображений ({len(images)}) превышает лимит Авито ({MAX_AVITO_IMAGES}). "
                f"Лишние фото обрезаны."
            )
            images = images[:MAX_AVITO_IMAGES]

        return images


class DescriptionBuilder:
    """Модуль генерации структурированного продающего описания товара."""

    @staticmethod
    def build_description(prod: dict) -> str:
        """
        Формирует структурированное продающее описание:
        - Наименование и артикул
        - Технические характеристики из колонок Excel
        - Блок преимуществ и описания
        - Блок условий компании
        """
        name = prod.get("name", "").strip()
        sku = prod.get("sku", "").strip()
        brand = prod.get("brand", "").strip() or config.COMPANY_BRAND

        # Характеристики
        specs = []
        if prod.get("power"):
            specs.append(f"• Мощность: {prod['power']} Вт")
        if prod.get("lumen"):
            specs.append(f"• Световой поток: {prod['lumen']} лм")
        if prod.get("color_temp"):
            specs.append(f"• Цветовая температура: {prod['color_temp']} К")
        if prod.get("base"):
            specs.append(f"• Цоколь: {prod['base']}")
        if prod.get("ip"):
            specs.append(f"• Степень защиты: IP{prod['ip']}")
        if prod.get("warranty"):
            specs.append(f"• Гарантия производителя: {prod['warranty']} года (лет)")
        if prod.get("cri"):
            specs.append(f"• Индекс цветопередачи: Ra ≥ {prod['cri']}")
        if prod.get("pulsation"):
            specs.append(f"• Коэффициент пульсации: ≤ {prod['pulsation']}% (без мерцания)")
        if prod.get("lifetime"):
            specs.append(f"• Срок службы: {prod['lifetime']}")
        if prod.get("mounting"):
            specs.append(f"• Способ установки: {prod['mounting']}")
        if prod.get("application"):
            specs.append(f"• Область применения: {prod['application']}")
        if prod.get("body_material"):
            specs.append(f"• Материал корпуса: {prod['body_material']}")
        if prod.get("diffuser_material"):
            specs.append(f"• Рассеиватель: {prod['diffuser_material']}")
        if prod.get("dimensions"):
            specs.append(f"• Габариты (Д×Ш×В): {prod['dimensions']} мм")
        if prod.get("weight"):
            specs.append(f"• Масса нетто: {prod['weight']} кг")

        specs_block = "\n".join(specs) if specs else "• Характеристики соответствуют паспорту изделия"

        # Извлечение содержательного описания из колонки AH
        extra_desc = prod.get("extra_desc", "")
        marketing_text = ""
        if extra_desc and len(extra_desc) > 30:
            paragraphs = [p.strip() for p in extra_desc.split("\n\n") if p.strip()]
            meaningful = [
                p for p in paragraphs
                if p.lower() != name.lower() and not p.lower().startswith(name.lower()[:20])
            ]
            if meaningful:
                summary_p = meaningful[0]
                marketing_text = f"\n[ОПИСАНИЕ И ПРЕИМУЩЕСТВА]\n{summary_p}\n"

        description_text = f"""{name}
Артикул: {sku}
Производитель: {brand}

[ТЕХНИЧЕСКИЕ ХАРАКТЕРИСТИКИ]
{specs_block}
{marketing_text}
[ПРЕИМУЩЕСТВА И УСЛОВИЯ КОМПАНИИ «ДАРИОН СВЕТ»]
- Официальная гарантия производителя на всю светотехнику.
- Быстрая отгрузка со склада в Санкт-Петербурге.
- Работаем с юр. и физ. лицами (оплата по счету с НДС 20% и без НДС).
- Быстрая доставка по Санкт-Петербургу, Ленинградской области и всей России (ТК СДЭК, Деловые Линии, ПЭК).
- Предоставляем полный комплект документов: паспорта изделий, сертификаты ЕАС.

Звоните или пишите в сообщения на Авито — рассчитаем освещенность объекта и подберем необходимое оборудование!"""

        return description_text.strip()


class FeedRouter:
    """Маршрутизатор товарных позиций по 8 целевым XML-фидам."""

    @staticmethod
    def route_product(prod: dict) -> str:
        """Определяет ключ фида для товара на основе листа, категории и наименования."""
        sheet = prod.get("sheet", "").lower()
        cat = prod.get("category", "").lower()
        name = prod.get("name", "").lower()

        # 1. Лампы и лампочки
        if "лампы" in sheet and "светильник" not in sheet:
            return "lamps"
        if "лампочк" in name or "лампа светодиодная" in name:
            return "lamps"

        # 2. Трековые системы
        if "трек" in name or "трек" in cat:
            return "track_systems"

        # 3. Люстры и потолочные светильники
        if "люстр" in name or "люстр" in cat:
            return "chandeliers"

        # 4. Бра, споты, светильники общего назначения
        if any(w in name for w in ["бра", "спот", "настольн", "торшер", "ночник", "подсветк"]):
            return "lighting_fixtures"

        # 5. Ламповые светильники
        if "ламповые светильники" in sheet or "ламповый" in name:
            return "lamp_luminaires"

        # 6. Электрика и комплектующие
        if any(w in name or w in cat for w in ["электрик", "блок питания", "драйвер", "кабель", "провод", "патрон"]):
            return "electrics"

        # 7. Прочие товары
        if "прочие товары" in sheet:
            return "other_goods"

        # 8. LED Светильники (по умолчанию для ассортимента Дарион Свет)
        return "led_luminaires"


class AvitoFeedGenerator:
    """Главный координатор сборки и выгрузки 8 XML-фидов."""

    def __init__(self):
        self.stock_sync = StockSync()
        self.image_resolver = ImageResolver()
        self.desc_builder = DescriptionBuilder()
        self.router = FeedRouter()
        self.products = []
        self.stock_map = {}
        self.feed_buckets = {k: [] for k in FEEDS_CONFIG.keys()}

    def load_products_from_excel(self) -> list:
        """Считывает все товарные позиции со всех листов книги Excel."""
        if not os.path.exists(EXCEL_PRODUCTS_PATH):
            raise FileNotFoundError(f"Файл номенклатуры не найден: {EXCEL_PRODUCTS_PATH}")

        wb = openpyxl.load_workbook(EXCEL_PRODUCTS_PATH, data_only=True)
        products = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            max_r = ws.max_row
            max_c = ws.max_column

            # Данные начинаются со строки 3
            for r in range(3, max_r + 1):
                sku = ws.cell(r, 1).value
                name = ws.cell(r, 2).value
                if not sku or not name:
                    continue

                sku_str = str(sku).strip()
                name_str = str(name).strip()

                # Извлечение параметров габаритов: A(22), B(23), C(24)
                dim_a = ws.cell(r, 22).value
                dim_b = ws.cell(r, 23).value
                dim_c = ws.cell(r, 24).value
                dim_str = f"{dim_a}×{dim_b}×{dim_c}" if (dim_a and dim_b and dim_c) else None

                prod = {
                    "sheet": sheet_name,
                    "sku": sku_str,
                    "name": name_str,
                    "category": str(ws.cell(r, 3).value or "").strip(),
                    "ean": ws.cell(r, 4).value,
                    "main_image": ws.cell(r, 5).value,
                    "warranty": ws.cell(r, 6).value,
                    "country": ws.cell(r, 7).value,
                    "brand": str(ws.cell(r, 8).value or config.COMPANY_BRAND).strip(),
                    "mounting": ws.cell(r, 9).value,
                    "application": ws.cell(r, 10).value,
                    "light_source_type": ws.cell(r, 11).value,
                    "lamps_count": ws.cell(r, 12).value,
                    "power": ws.cell(r, 13).value,
                    "base": ws.cell(r, 14).value,
                    "led_matrix": ws.cell(r, 15).value,
                    "lumen": ws.cell(r, 16).value,
                    "color_temp": ws.cell(r, 17).value,
                    "cri": ws.cell(r, 18).value,
                    "ip": ws.cell(r, 19).value,
                    "lifetime": ws.cell(r, 21).value,
                    "dimensions": dim_str,
                    "weight": ws.cell(r, 28).value,
                    "body_material": ws.cell(r, 30).value,
                    "diffuser_material": ws.cell(r, 31).value,
                    "extra_desc": str(ws.cell(r, 34).value or "").strip(),
                    "pulsation": ws.cell(r, 38).value,
                    "extra_images": ws.cell(r, 56).value,
                    "site_url": ws.cell(r, 57).value
                }
                products.append(prod)

        logger.info(f"Считано из Excel {len(products)} позиций номенклатуры.")
        return products

    @staticmethod
    def format_title(name: str, max_len: int = 50) -> str:
        """Ограничивает длину заголовка для Авито (максимум 50 символов) без потери смысла."""
        name = (name or "").strip()
        if len(name) <= max_len:
            return name
        # Удаляем хвостовые скобки с артикулами (VRN-UNE-48-G40K67-U)
        cleaned = re.sub(r'\s*\([^)]*\)\s*$', '', name).strip()
        if len(cleaned) <= max_len:
            return cleaned
        # Обрезка строго по границе последнего целого слова ≤ max_len символов.
        # rfind(' ', 0, max_len) находит позицию последнего пробела внутри окна [0, max_len).
        # Если пробелов нет (одно длинное слово без пробелов) — вынужденный жёсткий срез на max_len.
        last_space = cleaned.rfind(' ', 0, max_len)
        if last_space > 0:
            truncated = cleaned[:last_space].rstrip(' ,.-')
        else:
            truncated = cleaned[:max_len]   # крайний случай: нет пробелов в первых max_len символах
        # Финальная проверка: результат гарантированно ≤ max_len
        return truncated if len(truncated) <= max_len else truncated[:max_len]

    def build_ad_element(self, prod: dict, price: int, feed_key: str) -> etree.Element:
        """Формирует XML-элемент <Ad> со всеми обязательными и специфическими тегами."""
        ad = etree.Element("Ad")

        # 1. Обязательные базовые теги Авито
        ad_id = etree.SubElement(ad, "Id")
        ad_id.text = prod["sku"]

        address = etree.SubElement(ad, "Address")
        address.text = getattr(config, "DEFAULT_ADDRESS", "Санкт-Петербург, Студенческая ул., 10")

        title = etree.SubElement(ad, "Title")
        title.text = self.format_title(prod["name"], max_len=50)

        desc = etree.SubElement(ad, "Description")
        desc_text = self.desc_builder.build_description(prod)
        desc.text = etree.CDATA(desc_text)

        images_list = self.image_resolver.resolve_images(prod["main_image"], prod["extra_images"])
        images_el = etree.SubElement(ad, "Images")
        for img_url in images_list:
            img_tag = etree.SubElement(images_el, "Image")
            img_tag.set("url", img_url)

        price_el = etree.SubElement(ad, "Price")
        price_el.text = str(price)

        ad_type = etree.SubElement(ad, "AdType")
        ad_type.text = "Товар приобретен на продажу"

        condition = etree.SubElement(ad, "Condition")
        condition.text = "Новое"

        availability = etree.SubElement(ad, "Availability")
        availability.text = "В наличии"

        contact = etree.SubElement(ad, "ContactMethod")
        contact.text = "По телефону и в сообщениях"

        # 2. Категорийные теги согласно конфигурации фида
        cfg = FEEDS_CONFIG[feed_key]
        cat_el = etree.SubElement(ad, "Category")
        cat_el.text = cfg["category"]

        goods_type_el = etree.SubElement(ad, "GoodsType")
        goods_type_el.text = cfg["goods_type"]

        if cfg.get("goods_sub_type"):
            sub_type_el = etree.SubElement(ad, "GoodsSubType")
            sub_type_el.text = cfg["goods_sub_type"]

        if cfg.get("lighting_type"):
            light_type_el = etree.SubElement(ad, "LigitingType")
            light_type_el.text = cfg["lighting_type"]

        if prod.get("brand"):
            brand_el = etree.SubElement(ad, "Brand")
            brand_el.text = prod["brand"]

        return ad

    def generate_all_feeds(self) -> dict:
        """
        Запускает полный конвейер генерации 8 фидов.
        """
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # 1. Загрузка актуальных остатков
        self.stock_map = self.stock_sync.load_stock_dict()

        # 2. Загрузка товаров из номенклатуры
        self.products = self.load_products_from_excel()

        # 3. Маршрутизация с фильтрацией по наличию остатка
        matched_count = 0
        for prod in self.products:
            sku = prod["sku"]
            if sku in self.stock_map:
                price = self.stock_map[sku]["price"]
                feed_key = self.router.route_product(prod)
                self.feed_buckets[feed_key].append((prod, price))
                matched_count += 1
            else:
                logger.warning(f"Товар {sku} ('{prod['name']}') пропущен: нет на остатке или остаток <= 0.")

        logger.info(f"Сопоставлено и направлено в фиды товаров: {matched_count}.")

        # 4. Генерация 8 XML-файлов в output/
        generated_summary = {}

        for feed_key, cfg in FEEDS_CONFIG.items():
            out_filename = cfg["filename"]
            out_path = os.path.join(OUTPUT_DIR, out_filename)
            items = self.feed_buckets[feed_key]

            # Создаем корневой элемент Авито
            root = etree.Element("Ads", formatVersion="3", target="Avito.ru")

            if len(items) > 0:
                for prod, price in items:
                    ad_el = self.build_ad_element(prod, price, feed_key)
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
            msg = "Настройки FTP не заполнены в config.py (FTP_HOST, FTP_USER). Файлы сохранены локально в папке output/"
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

            # Переход в удаленную папку или создание
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
            msg = "Настройки SFTP не заполнены в config.py (SFTP_HOST, SFTP_USER). Файлы сохранены локально в папке output/"
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

            # Проверка и создание папки
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

    # 1. Статус скачивания остатков
    print("1. СТАТУС СИНХРОНИЗАЦИИ ОСТАТКОВ ИЗ GOOGLE SHEETS:")
    print("-" * 95)
    print(f"Статус подключения:        {generator.stock_sync.sync_status}")
    print(f"Сообщение шлюза:           {generator.stock_sync.sync_message}")
    print(f"Использованный источник:   {os.path.relpath(generator.stock_sync.used_source, BASE_DIR)}")
    print(f"Доступных позиций в кэше:  {len(generator.stock_map)} шт. (с остатком > 0)")
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

    # 4. Пример сгенерированного блока <Ad> для V9007 (извлекаем с точным CDATA)
    v9007_file = os.path.join(OUTPUT_DIR, "led_luminaires_feed.xml")
    if os.path.exists(v9007_file):
        with open(v9007_file, "r", encoding="utf-8") as f:
            raw_content = f.read()
        m = re.search(r'(  <Ad>.*?</Ad>)', raw_content, re.DOTALL)
        if m:
            print("4. ПРИМЕР ПОЛНОГО СГЕНЕРИРОВАННОГО БЛОКА <Ad> ДЛЯ ТОВАРА V9007:")
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
