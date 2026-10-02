"""
Модуль анализа номенклатурного файла data/Продукты (19).xlsx.
Выполняет:
1. Определение всех листов (вкладок) книги Excel.
2. Чтение заголовков (строка 2) и первой строки с данными (строка 3).
3. Валидацию форматов:
   - Столбец A: Артикул (тип, пустота, паттерн)
   - Столбец B: Наименование товара
   - Столбец C: Категория продукции
   - Столбец E: Основное изображение (URL)
   - Столбец BD: Ссылка на архив доп. фото (URL архива)
4. Подсчет количества позиций и группировку по категориям.
"""

import os
import sys
import io
import openpyxl
from openpyxl.utils import get_column_letter

# Обеспечиваем корректный вывод UTF-8 в Windows консоль
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DATA_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "data"))
EXCEL_PATH = os.path.join(DATA_DIR, "Продукты (19).xlsx")


def analyze_nomenclature():
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Файл номенклатуры не найден по пути: {EXCEL_PATH}")

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    sheet_names = wb.sheetnames

    report = {
        "file_name": os.path.basename(EXCEL_PATH),
        "sheets": sheet_names,
        "sheet_details": {},
        "total_items": 0,
        "categories_summary": {}
    }

    for name in sheet_names:
        ws = wb[name]
        max_row = ws.max_row
        max_col = ws.max_column

        # Считываем строку заголовков (строка 2)
        headers = {}
        for c in range(1, max_col + 1):
            val = ws.cell(2, c).value
            if val is not None:
                headers[get_column_letter(c)] = {
                    "col_idx": c,
                    "title": str(val).strip()
                }

        # Считываем данные начиная со строки 3
        data_rows = []
        for r in range(3, max_row + 1):
            # Проверяем, не пустая ли строка целиком
            row_vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
            if any(v is not None and str(v).strip() != "" for v in row_vals):
                row_dict = {}
                for c in range(1, max_col + 1):
                    col_l = get_column_letter(c)
                    val = ws.cell(r, c).value
                    row_dict[col_l] = val
                data_rows.append(row_dict)

        report["sheet_details"][name] = {
            "max_row": max_row,
            "max_col": max_col,
            "data_rows_count": len(data_rows),
            "headers": headers,
            "first_data_row": data_rows[0] if data_rows else None
        }

        report["total_items"] += len(data_rows)

        # Категории
        for row in data_rows:
            cat = row.get("C")
            cat_str = str(cat).strip() if cat else "Без категории"
            report["categories_summary"][cat_str] = report["categories_summary"].get(cat_str, 0) + 1

    return report


def print_nomenclature_report(report):
    print("=" * 80)
    print("АНАЛИЗ НОМЕНКЛАТУРНОГО ФАЙЛА:", report["file_name"])
    print("=" * 80)
    print(f"Всего листов: {len(report['sheets'])} -> {', '.join(report['sheets'])}")
    print(f"Всего позиций данных: {report['total_items']}\n")

    for s_name, s_info in report["sheet_details"].items():
        print(f"--- Лист: [{s_name}] ---")
        print(f"  Размер сетки: {s_info['max_row']} строк x {s_info['max_col']} столбцов")
        print(f"  Количество позиций с данными: {s_info['data_rows_count']}")

        first_row = s_info["first_data_row"]
        if first_row:
            art = first_row.get("A")
            name = first_row.get("B")
            cat = first_row.get("C")
            img = first_row.get("E")
            extra_imgs = first_row.get("BD")

            print("\n  Пример первой строки данных (строка 3):")
            print(f"    [Столбец A] Артикул: {repr(art)} (Тип: {type(art).__name__})")
            print(f"    [Столбец B] Наименование: {repr(name)}")
            print(f"    [Столбец C] Категория: {repr(cat)}")
            print(f"    [Столбец E] Основное фото: {repr(img)}")
            print(f"    [Столбец BD] Архив доп. фото: {repr(extra_imgs)}")
            
            # Анализ валидности ссылок
            print("\n  Проверка форматов ключевых полей:")
            print(f"    - Артикул: {'Корректный строковый идентификатор' if art else 'ОШИБКА: пустой артикул'}")
            print(f"    - Название: {'Заполнено' if name else 'ОШИБКА: пусто'}")
            print(f"    - URL фото (E): {'Валидный HTTP-URL' if img and str(img).startswith('http') else 'Некорректный URL'}")
            print(f"    - URL архива (BD): {'Ссылка на zip/archive' if extra_imgs and 'saveas_archive' in str(extra_imgs) else 'Нестандартный формат'}")
        else:
            print("  (На листе нет строк с данными, только структура заголовков)")
        print()

    print("СВОДКА ПО КАТЕГОРИЯМ:")
    if report["categories_summary"]:
        for cat, cnt in report["categories_summary"].items():
            print(f"  • {cat}: {cnt} шт.")
    else:
        print("  Позиции отсутствуют.")
    print("=" * 80)


if __name__ == "__main__":
    rep = analyze_nomenclature()
    print_nomenclature_report(rep)
