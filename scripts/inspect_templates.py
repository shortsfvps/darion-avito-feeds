"""
Модуль разведки и структурирования шаблонов Авито (Автозагрузка).
Выполняет:
1. Чтение каждого XML-файла в папке templates/.
2. Извлечение тегов <GoodsSubType>, <Category>, <Title> или структурный анализ специфических тегов категорий Авито.
3. Определение понятного человекочитаемого имени шаблона.
4. Переименование файлов f_*.xml в понятные имена (track_systems.xml, chandeliers.xml и т.д.).
5. Сохранение схемы тегов каждого шаблона в data/template_tags_schema.json.
"""

import os
import sys
import io
import glob
import json
import shutil
import xml.etree.ElementTree as ET

# Обеспечиваем корректный вывод UTF-8 в Windows консоль
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

TEMPLATES_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "templates"))
DATA_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "data"))
SCHEMA_FILE = os.path.join(DATA_DIR, "template_tags_schema.json")
MANIFEST_FILE = os.path.join(TEMPLATES_DIR, "templates_manifest.json")

# Известные правила классификации шаблонов на основе специфических тегов Авито
SIGNATURE_MAP = {
    # f_8046: Трековые системы
    "track_systems": {
        "required_tags": ["TrackLightingSystemType", "InstallationMethod", "StretchCeiling"],
        "target_filename": "track_systems.xml",
        "category_title": "Трековые системы освещения",
        "category": "Для дома и дачи / Мебель и интерьер / Освещение",
        "goods_sub_type": "Освещение",
        "lighting_type": "Трековые системы"
    },
    # f_2536: Люстры и потолочные светильники
    "chandeliers": {
        "required_tags": ["ChandelierType", "ChandelierMountingType", "LedLamp"],
        "target_filename": "chandeliers.xml",
        "category_title": "Люстры и потолочные светильники",
        "category": "Для дома и дачи / Мебель и интерьер / Освещение",
        "goods_sub_type": "Освещение",
        "lighting_type": "Потолочные светильники"
    },
    # f_3456: Лампы и лампочки
    "lamps": {
        "required_tags": ["BulbType", "BulbBaseType", "GlowColor", "Temperature", "Power"],
        "target_filename": "lamps.xml",
        "category_title": "Лампы / Лампочки",
        "category": "Для дома и дачи / Мебель и интерьер / Освещение",
        "goods_sub_type": "Освещение",
        "lighting_type": "Лампочки"
    },
    # f_2216: Электрика и комплектующие
    "electrics": {
        "required_tags": ["ElectricsType"],
        "target_filename": "electrics.xml",
        "category_title": "Электрика и комплектующие",
        "category": "Ремонт и строительство / Электрика",
        "goods_sub_type": "Электрика",
        "lighting_type": None
    },
    # f_6746: Светильники общего назначения / бра / споты
    "lighting_fixtures": {
        "required_tags": ["LigitingType"],
        "exclude_tags": ["TrackLightingSystemType", "ChandelierType", "BulbType"],
        "target_filename": "lighting_fixtures.xml",
        "category_title": "Светильники общего назначения (Бра, Споты, Уличные)",
        "category": "Для дома и дачи / Мебель и интерьер / Освещение",
        "goods_sub_type": "Освещение",
        "lighting_type": "Светильники"
    },
    # f_2996: Прочие товары (базовый шаблон без GoodsSubType / LigitingType)
    "other_goods": {
        "required_tags": [],
        "exclude_tags": ["GoodsSubType", "LigitingType", "ElectricsType"],
        "target_filename": "other_goods.xml",
        "category_title": "Прочие товары / Торговое оборудование",
        "category": "Для дома и дачи / Оборудование для бизнеса",
        "goods_sub_type": None,
        "lighting_type": None
    }
}

# Для шаблонов f_5346 и f_5716 (оба содержат GoodsSubType, но не LigitingType)
# Они соответствуют оставшимся категориям номенклатуры: LED Светильники и Ламповые светильники
PAIR_MAPPING = {
    "f_5346abf5b7119f2d.xml": {
        "target_filename": "led_luminaires.xml",
        "category_title": "LED Светильники (промышленные, консольные, универсальные)",
        "category": "Для дома и дачи / Мебель и интерьер / Освещение",
        "goods_sub_type": "Светильники LED"
    },
    "f_5716abf5b71299ac.xml": {
        "target_filename": "lamp_luminaires.xml",
        "category_title": "Ламповые светильники",
        "category": "Для дома и дачи / Мебель и интерьер / Освещение",
        "goods_sub_type": "Светильники ламповые"
    }
}


def analyze_template_file(file_path):
    """Считывает XML-файл, извлекает теги, текст значений и определяет категорию."""
    tree = ET.parse(file_path)
    root = tree.getroot()
    ad = root.find("Ad")
    if ad is None:
        tags = []
        tag_values = {}
    else:
        tags = [child.tag for child in ad]
        tag_values = {
            child.tag: (child.text.strip() if child.text and child.text.strip() else "")
            for child in ad
        }

    # Пробуем извлечь явные значения тегов
    goods_sub_type = tag_values.get("GoodsSubType", "")
    category = tag_values.get("Category", "")
    title = tag_values.get("Title", "")
    goods_type = tag_values.get("GoodsType", "")
    lighting_type = tag_values.get("LigitingType", "")

    filename = os.path.basename(file_path)

    # Загружаем манифест при наличии
    manifest = {}
    if os.path.exists(MANIFEST_FILE):
        try:
            with open(MANIFEST_FILE, "r", encoding="utf-8") as mf:
                manifest = json.load(mf)
        except Exception:
            manifest = {}

    orig_name = manifest.get(filename, {}).get("original_name", filename)

    # Идентификация по сигнатуре
    identified_key = None
    target_filename = None
    category_title = None

    # Проверка пары f_5346 / f_5716 (или их переименованных версий)
    if orig_name in PAIR_MAPPING:
        target_filename = PAIR_MAPPING[orig_name]["target_filename"]
        category_title = PAIR_MAPPING[orig_name]["category_title"]
        identified_key = target_filename.replace(".xml", "")
    elif filename == "led_luminaires.xml":
        target_filename = "led_luminaires.xml"
        category_title = "LED Светильники (промышленные, консольные, универсальные)"
    elif filename == "lamp_luminaires.xml":
        target_filename = "lamp_luminaires.xml"
        category_title = "Ламповые светильники"
    else:
        # Проверка по тегам сигнатур
        for key, sig in SIGNATURE_MAP.items():
            req = sig["required_tags"]
            exc = sig.get("exclude_tags", [])
            has_req = all(t in tags for t in req) if req else True
            has_no_exc = not any(t in tags for t in exc)
            if has_req and has_no_exc:
                identified_key = key
                target_filename = sig["target_filename"]
                category_title = sig["category_title"]
                break

    if not target_filename:
        target_filename = filename
        category_title = filename

    return {
        "original_file": orig_name,
        "current_file": filename,
        "full_path": file_path,
        "target_filename": target_filename,
        "category_title": category_title,
        "tags_count": len(tags),
        "tags": tags,
        "extracted_values": {
            "GoodsSubType": goods_sub_type,
            "Category": category,
            "GoodsType": goods_type,
            "Title": title,
            "LigitingType": lighting_type
        }
    }


def inspect_and_rename_templates():
    os.makedirs(DATA_DIR, exist_ok=True)
    xml_files = sorted(glob.glob(os.path.join(TEMPLATES_DIR, "*.xml")))

    # Исключаем временные или уже сохраненные не-шаблонные файлы
    template_files = [f for f in xml_files if not os.path.basename(f).startswith("temp_")]

    results = []
    rename_plan = []

    for fpath in template_files:
        info = analyze_template_file(fpath)
        results.append(info)
        curr = os.path.basename(fpath)
        target = info["target_filename"]
        orig = info["original_file"]
        if curr != target:
            rename_plan.append((fpath, os.path.join(TEMPLATES_DIR, target), orig, target))

    # Выполняем переименование (с сохранением резервной копии оригиналов в manifest)
    manifest = {}
    if os.path.exists(MANIFEST_FILE):
        try:
            with open(MANIFEST_FILE, "r", encoding="utf-8") as mf:
                manifest = json.load(mf)
        except Exception:
            manifest = {}

    for src_path, dst_path, orig_name, target_name in rename_plan:
        if src_path != dst_path:
            shutil.copy2(src_path, dst_path)
            os.remove(src_path)
        manifest[target_name] = {
            "original_name": orig_name,
            "new_name": target_name
        }

    # Сохраняем манифест
    if rename_plan or not os.path.exists(MANIFEST_FILE):
        with open(MANIFEST_FILE, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)

    schema_data = {}
    for r in results:
        target_name = r["target_filename"]
        schema_data[target_name] = {
            "category_title": r["category_title"],
            "original_filename": r["original_file"],
            "total_tags": r["tags_count"],
            "tags_list": r["tags"],
            "extracted_tags": r["extracted_values"]
        }

    with open(SCHEMA_FILE, "w", encoding="utf-8") as f:
        json.dump(schema_data, f, ensure_ascii=False, indent=2)

    return results, rename_plan


if __name__ == "__main__":
    results, renames = inspect_and_rename_templates()
    print("=" * 80)
    print("ОТЧЕТ ПО ШАБЛОНАМ АВИТО (TEMPLATES)")
    print("=" * 80)
    for r in results:
        print(f"Файл: {r['original_file']} -> {r['target_filename']}")
        print(f"  Назначение: {r['category_title']}")
        print(f"  Количество тегов: {r['tags_count']}")
        key_tags = [t for t in r['tags'] if t in [
            'ElectricsType', 'ChandelierType', 'ChandelierMountingType',
            'TrackLightingSystemType', 'BulbType', 'BulbBaseType', 'LigitingType', 'GoodsSubType'
        ]]
        print(f"  Специфические теги: {', '.join(key_tags) if key_tags else 'Базовый набор'}")
        print("-" * 80)
    print(f"\nВсего обработано: {len(results)} шаблонов.")
    print(f"Переименовано файлов: {len(renames)}.")
    print(f"Схема тегов сохранена в: {SCHEMA_FILE}")
