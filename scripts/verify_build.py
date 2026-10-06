"""
Скрипт сквозного аудита и верификации сборки (QA / DevOps).
Проверяет:
1. Конфигурацию и окружение (config.py, requirements.txt, run.bat).
2. Исполнение генератора (штатный запуск и сценарий отката на кэш при сетевой ошибке).
3. Валидацию сгенерированных XML-фидов в output/ (strict XML parsing, теги, UTF-8, отсутствие &amp;amp;, длина Title <= 50).
4. Наличие и структуру пользовательской документации (docs/ИНСТРУКЦИЯ_ПОЛЬЗОВАТЕЛЯ.md).
"""

import os
import sys
import io
import re
import importlib
import requests
from lxml import etree

# Настройка безопасного UTF-8 вывода в консоль Windows
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


class BuildVerifier:
    def __init__(self):
        self.results = []

    def check(self, category: str, name: str, condition: bool, details: str = ""):
        status = "PASS" if condition else "FAIL"
        self.results.append({
            "category": category,
            "name": name,
            "status": status,
            "details": details
        })
        mark = "✓ [PASS]" if condition else "✗ [FAIL]"
        print(f"  {mark} {name}: {details}")

    def verify_all(self):
        print("\n" + "=" * 95)
        print("          СКВОЗНОЙ АУДИТ И ВЕРИФИКАЦИЯ СБОРКИ (QA / DEVOPS АТТЕСТАЦИЯ)")
        print("=" * 95 + "\n")

        # ----------------------------------------------------------------------
        # 1. КОНФИГУРАЦИЯ И ОКРУЖЕНИЕ
        # ----------------------------------------------------------------------
        print("1. ПРОВЕРКА КОНФИГУРАЦИИ И ОКРУЖЕНИЯ:")
        config_path = os.path.join(BASE_DIR, "config.py")
        has_config = os.path.exists(config_path)
        self.check("Environment", "Наличие config.py", has_config, f"Путь: {config_path}")

        config_valid = False
        upload_method_safe = False
        if has_config:
            try:
                import config
                importlib.reload(config)
                config_valid = True
                upload_method_safe = (getattr(config, "UPLOAD_METHOD", None) == "local")
            except Exception as e:
                config_valid = False
                upload_method_safe = False

        self.check(
            "Environment",
            "Чтение config.py без синтаксических ошибок",
            config_valid,
            "Модуль config успешно импортирован" if config_valid else "Ошибка импорта config"
        )
        self.check(
            "Environment",
            "Безопасный режим публикации по умолчанию (UPLOAD_METHOD='local')",
            upload_method_safe,
            f"Текущее значение: UPLOAD_METHOD = '{getattr(config, 'UPLOAD_METHOD', 'N/A')}'"
        )

        # requirements.txt
        req_path = os.path.join(BASE_DIR, "requirements.txt")
        has_req = os.path.exists(req_path)
        req_modules = []
        if has_req:
            with open(req_path, "r", encoding="utf-8") as f:
                req_text = f.read()
                req_modules = [line.strip().split("==")[0].lower() for line in req_text.splitlines() if line.strip()]

        required_libs = ["pandas", "openpyxl", "requests", "lxml", "paramiko"]
        all_libs_present = all(lib in req_modules for lib in required_libs)
        self.check(
            "Environment",
            "Состав зависимостей в requirements.txt",
            all_libs_present,
            f"Зафиксированы: {', '.join(required_libs)}"
        )

        # run.bat
        bat_path = os.path.join(BASE_DIR, "run.bat")
        has_bat = os.path.exists(bat_path)
        bat_has_chcp = False
        bat_has_where_python = False
        bat_has_errorlevel = False

        if has_bat:
            with open(bat_path, "r", encoding="utf-8", errors="ignore") as f:
                bat_content = f.read()
                bat_has_chcp = "chcp 65001" in bat_content
                bat_has_where_python = "where python" in bat_content
                bat_has_errorlevel = "errorlevel" in bat_content.lower()

        self.check("Environment", "Наличие исполняемого файла run.bat", has_bat, f"Путь: {bat_path}")
        self.check("Environment", "Кодировка консоли chcp 65001 в run.bat", bat_has_chcp, "UTF-8 кодировка объявлена")
        self.check("Environment", "Проверка наличия Python в run.bat", bat_has_where_python, "Команда where python присутствует")
        self.check("Environment", "Обработка кодов возврата errorlevel в run.bat", bat_has_errorlevel, "Контроль ошибок execution присутствует")
        print()

        # ----------------------------------------------------------------------
        # 2. ИСПОЛНЕНИЕ ГЕНЕРАТОРА И ОТКАЗОУСТОЙЧИВОСТЬ
        # ----------------------------------------------------------------------
        print("2. ТЕСТИРОВАНИЕ ИСПОЛНЕНИЯ ГЕНЕРАТОРА И ОТКАЗОУСТОЙЧИВОСТИ:")
        from scripts.generate_feeds import StockSync, AvitoFeedGenerator, FeedUploader

        # Тест штатной синхронизации
        stock_sync = StockSync()
        stock_dict = stock_sync.load_stock_dict()
        is_stock_loaded = len(stock_dict) > 0
        self.check(
            "Generator",
            "Загрузка остатков и парсинг (онлайн Google Sheets / кэш)",
            is_stock_loaded,
            f"Статус: {stock_sync.sync_status}, загружено {len(stock_dict)} позиций с остатком > 0"
        )

        # Тест имитации сбоя сети (проверка отката на резервный кэш remains_spb_template.xlsx)
        fallback_sync = StockSync(
            url="https://invalid-non-existing-domain-test-12345.com/export.xlsx",
            latest_cache=os.path.join(BASE_DIR, "data", "non_existing_cache_test.xlsx"),
            fallback_cache=os.path.join(BASE_DIR, "data", "remains_spb_template.xlsx")
        )
        fallback_data = fallback_sync.load_stock_dict()
        fallback_success = (len(fallback_data) > 0 and "V9007" in fallback_data)
        self.check(
            "Generator",
            "Отказоустойчивость: переключение на remains_spb_template.xlsx при недоступности сети",
            fallback_success,
            f"Резервный кэш успешно прочитан: {len(fallback_data)} позиций, V9007={fallback_data.get('V9007')}"
        )

        # Программный запуск генерации
        gen = AvitoFeedGenerator()
        summary = gen.generate_all_feeds()
        gen_success = (len(summary) == 8)
        self.check(
            "Generator",
            "Программное выполнение generate_all_feeds() в режиме local",
            gen_success,
            f"Сформировано фидов: {len(summary)} из 8"
        )
        print()

        # ----------------------------------------------------------------------
        # 3. ВАЛИДАЦИЯ СГЕНЕРИРОВАННЫХ АРТЕФАКТОВ В output/
        # ----------------------------------------------------------------------
        print("3. СТРОГАЯ ВАЛИДАЦИЯ АРТЕФАКТОВ (output/):")
        output_dir = os.path.join(BASE_DIR, "output")
        expected_feeds = [
            "led_luminaires_feed.xml",
            "track_systems_feed.xml",
            "chandeliers_feed.xml",
            "lamps_feed.xml",
            "lighting_fixtures_feed.xml",
            "lamp_luminaires_feed.xml",
            "electrics_feed.xml",
            "other_goods_feed.xml"
        ]

        all_files_exist = all(os.path.exists(os.path.join(output_dir, f)) for f in expected_feeds)
        self.check(
            "Artifacts",
            "Наличие всех 8 целевых XML-фидов в папке output/",
            all_files_exist,
            f"Проверено 8 файлов: {', '.join(expected_feeds)}"
        )

        # Строгий парсинг без recover
        strict_parser = etree.XMLParser(recover=False, encoding="utf-8")
        all_parsed_strict = True
        encoding_declared_utf8 = True
        empty_feeds_valid = True

        for fname in expected_feeds:
            fpath = os.path.join(output_dir, fname)
            if not os.path.exists(fpath):
                all_parsed_strict = False
                continue

            with open(fpath, "rb") as fp:
                raw_bytes = fp.read()

            # Проверка объявления UTF-8 в заголовке
            if b'encoding="UTF-8"' not in raw_bytes and b"encoding='UTF-8'" not in raw_bytes and b'encoding="utf-8"' not in raw_bytes and b"encoding='utf-8'" not in raw_bytes:
                encoding_declared_utf8 = False

            try:
                tree = etree.fromstring(raw_bytes, parser=strict_parser)
                if tree.tag != "Ads":
                    all_parsed_strict = False
                if tree.attrib.get("formatVersion") != "3" or tree.attrib.get("target") != "Avito.ru":
                    all_parsed_strict = False

                # Проверка каркаса для пустых фидов
                if fname not in ["led_luminaires_feed.xml", "track_systems_feed.xml", "lamps_feed.xml"]:
                    ads = tree.findall("Ad")
                    if len(ads) != 0:
                        empty_feeds_valid = False
            except Exception as e:
                all_parsed_strict = False
                print(f"      [ОШИБКА ПАРСИНГА {fname}]: {e}")

        self.check(
            "Artifacts",
            "Строгий парсинг XML через lxml.etree.XMLParser(recover=False)",
            all_parsed_strict,
            "Все 8 файлов синтаксически корректны (валидные XML DOM-деревья)"
        )
        self.check(
            "Artifacts",
            "Корректная декларация кодировки UTF-8 в XML-прологе",
            encoding_declared_utf8,
            "XML декларация содержит encoding='UTF-8'"
        )
        self.check(
            "Artifacts",
            "Валидный каркас пустых фидов <Ads formatVersion='3' target='Avito.ru'></Ads>",
            empty_feeds_valid,
            "Все 5 пустых фидов содержат ровно 0 объявлений и валидный корневой тег"
        )

        # Глубокая валидация led_luminaires_feed.xml
        led_feed_path = os.path.join(output_dir, "led_luminaires_feed.xml")
        has_required_tags = False
        no_double_escaping = False
        title_length_valid = False
        title_val = ""
        title_len = 0

        if os.path.exists(led_feed_path):
            with open(led_feed_path, "r", encoding="utf-8") as fp:
                led_text = fp.read()

            # Проверка отсутствия двойного экранирования &amp;amp;
            no_double_escaping = ("&amp;amp;" not in led_text)

            tree = etree.parse(led_feed_path, parser=strict_parser)
            ad = tree.find("Ad")
            if ad is not None:
                id_tag = ad.find("Id")
                title_tag = ad.find("Title")
                price_tag = ad.find("Price")
                images_tag = ad.find("Images")
                desc_tag = ad.find("Description")

                has_required_tags = all(
                    t is not None and t.text and t.text.strip()
                    for t in [id_tag, title_tag, price_tag, desc_tag]
                ) and (images_tag is not None and len(images_tag.findall("Image")) > 0)

                if title_tag is not None and title_tag.text:
                    title_val = title_tag.text.strip()
                    title_len = len(title_val)
                    title_length_valid = (title_len <= 50)

        self.check(
            "Artifacts",
            "Наличие обязательных тегов Авито в led_luminaires_feed.xml (<Id>, <Title>, <Price>, <Images>, <Description>)",
            has_required_tags,
            "Все базовые теги присутствуют и непустые"
        )
        self.check(
            "Artifacts",
            "Отсутствие двойного экранирования амперсандов (&amp;amp;)",
            no_double_escaping,
            "Дублирующее экранирование отсутствует, ссылки на изображения валидны"
        )
        self.check(
            "Artifacts",
            f"Ограничение длины <Title> <= 50 символов (фактическая длина: {title_len})",
            title_length_valid,
            f"Заголовок: '{title_val}'"
        )

        # Глубокая валидация lamps_feed.xml
        lamps_feed_path = os.path.join(output_dir, "lamps_feed.xml")
        lamps_tags_valid = False
        lamps_no_obsolete_tags = False
        lamps_count = 0
        lamps_details = ""

        if os.path.exists(lamps_feed_path):
            tree = etree.parse(lamps_feed_path, parser=strict_parser)
            ads = tree.findall("Ad")
            lamps_count = len(ads)
            if lamps_count > 0:
                first_ad = ads[0]
                cat = first_ad.findtext("Category")
                gt = first_ad.findtext("GoodsType")
                gst = first_ad.findtext("GoodsSubType")
                lt = first_ad.findtext("LigitingType")
                bt = first_ad.findtext("BulbType")
                bbt = first_ad.findtext("BulbBaseType")
                br = first_ad.findtext("Brand")
                pw = first_ad.findtext("Power")
                bip = first_ad.findtext("BulbsInPackage")

                lamps_tags_valid = (
                    cat == "Мебель и интерьер" and
                    gt == "Освещение" and
                    gst == "Комплектующие" and
                    lt == "Лампочки" and
                    bt in ("Светодиодная", "Филаментная") and
                    bool(bbt) and
                    (pw and "Вт" in pw) and
                    bip == "1"
                )
                lamps_details = f"Объявлений: {lamps_count}, Category={cat}, GoodsSubType={gst}, LigitingType={lt}, BulbType={bt}"

                has_obsolete = any(
                    ad.find("LampType") is not None or
                    ad.find("CapType") is not None or
                    ad.find("ColorTemperature") is not None or
                    ad.find("LightFlux") is not None or
                    ad.findtext("Brand") == "RSV"
                    for ad in ads
                )
                lamps_no_obsolete_tags = not has_obsolete

                # Проверка фильтрации температур по белому списку Авито
                valid_temps_set = {
                    '1800 К', '2000 К', '2200 К', '2400 К', '2700 К', '2800 К',
                    '3000 К', '4000 К', '5000 К', '6000 К', '6400 К', '6500 К'
                }
                invalid_temps = [
                    ad.findtext("Temperature") for ad in ads
                    if ad.find("Temperature") is not None and ad.findtext("Temperature") not in valid_temps_set
                ]
                g_special_no_temp = all(
                    ad.find("Temperature") is None
                    for ad in ads if ad.findtext("Id") in ("G33127T", "G13629")
                )
                lamps_temp_valid = (len(invalid_temps) == 0 and g_special_no_temp)

        self.check(
            "Artifacts",
            "Соответствие схемы lamps_feed.xml (Category/GoodsType/GoodsSubType/LigitingType/BulbType/BulbBaseType/Power/BulbsInPackage)",
            lamps_tags_valid,
            lamps_details
        )
        self.check(
            "Artifacts",
            "Исключение недопустимых брендов (RSV) и устаревших тегов в lamps_feed.xml (<LampType>, <CapType>, <ColorTemperature>, <LightFlux>)",
            lamps_no_obsolete_tags,
            "Тег Brand=RSV и устаревшие теги отсутствуют во всех объявлениях фида ламп"
        )
        self.check(
            "Artifacts",
            "Фильтрация тега <Temperature> по словарю Авито (исключение нестандартных 4100 К / 4200 К на G33127T и G13629)",
            lamps_temp_valid,
            "Тег <Temperature> отсутствует на G33127T и G13629, все остальные значения входят в VALID_TEMPS"
        )
        print()

        # ----------------------------------------------------------------------
        # 4. НАЛИЧИЕ И ПОЛНОТА ДОКУМЕНТАЦИИ
        # ----------------------------------------------------------------------
        print("4. ПРОВЕРКА ДОКУМЕНТАЦИИ:")
        doc_path = os.path.join(BASE_DIR, "docs", "ИНСТРУКЦИЯ_ПОЛЬЗОВАТЕЛЯ.md")
        has_doc = os.path.exists(doc_path)
        doc_size = os.path.getsize(doc_path) if has_doc else 0

        has_sec1 = False
        has_sec2 = False
        has_sec3 = False
        has_sec4 = False

        if has_doc:
            with open(doc_path, "r", encoding="utf-8") as fp:
                doc_text = fp.read()
                has_sec1 = "Раздел 1" in doc_text and "8 XML-фидов" in doc_text
                has_sec2 = "Раздел 2" in doc_text and "run.bat" in doc_text
                has_sec3 = "Раздел 3" in doc_text and "Планировщик" in doc_text
                has_sec4 = "Раздел 4" in doc_text and "Автозагрузка" in doc_text

        all_sections_present = has_sec1 and has_sec2 and has_sec3 and has_sec4
        self.check("Documentation", "Наличие docs/ИНСТРУКЦИЯ_ПОЛЬЗОВАТЕЛЯ.md", has_doc, f"Размер: {doc_size} байт")
        self.check(
            "Documentation",
            "Полнота разделов инструкции (Разделы 1-4 включены)",
            all_sections_present,
            "Разделы: 1 (Состав фидов), 2 (Запуск run.bat), 3 (Task Scheduler), 4 (Кабинет Авито)"
        )
        print()

        # ----------------------------------------------------------------------
        # ИТОГОВЫЙ СТАТУС ВЕРИФИКАЦИИ
        # ----------------------------------------------------------------------
        total_checks = len(self.results)
        passed_checks = sum(1 for r in self.results if r["status"] == "PASS")
        failed_checks = total_checks - passed_checks

        print("=" * 95)
        print(f"ИТОГ АУДИТА: Успешно пройдено {passed_checks} из {total_checks} контрольных точек.")
        if failed_checks == 0:
            print("ВЕРДИКТ: [ СБОРКА ПОЛНОСТЬЮ ВЕРИФИЦИРОВАНА И ГОТОВА К ПЕРЕДАЧЕ ЗАКАЗЧИКУ ]")
        else:
            print(f"ВЕРДИКТ: [ ОБНАРУЖЕНО {failed_checks} ОШИБОК! ТРЕБУЕТСЯ ДОРАБОТКА ]")
        print("=" * 95 + "\n")

        return (failed_checks == 0)


if __name__ == "__main__":
    verifier = BuildVerifier()
    success = verifier.verify_all()
    sys.exit(0 if success else 1)
