import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import openpyxl

# Загрузка словаря соответствий из JSON
mapping_json = '''
{
  "Группа": {"ozon": "Название группы"},
  "Артикул продавца": {"ozon": "Артикул*"},
  "Артикул WB": {"ozon": "SKU"},
  "Наименование": {"ozon": "Название товара"},
  "Бренд": {"ozon": "Бренд в одежде и обуви*"},
  "Пол": {"ozon": "Пол*"},
  "Цвет": {"ozon": "Цвет товара*"},
  "Состав": {"ozon": "Состав материала"},
  "Материал подкладки": {"ozon": "Материал подкладки"},
  "Рост модели на фото": {"ozon": "Рост модели на фото"},
  "Параметры модели на фото (ОГ-ОТ-ОБ)": {"ozon": "Параметры модели на фото (ОГ-ОТ-ОБ)"},
  "Размер на модели": {"ozon": "Размер на модели"},
  "Коллекция": {"ozon": "Коллекция"},
  "Страна производства": {"ozon": "Страна-изготовитель"},
  "Рос. размер": {"ozon": "Российский размер*"},
  "Размер": {"ozon": "Размер производителя", "normalize": "concat_with|Рос. размер"},
  "Баркод": {"ozon": "Штрихкод (Серийный номер / EAN)", "aliases": ["Баркод", "Штрихкод (Серийный номер / EAN)", "Штрихкод", "EAN"]},
  "Вес с упаковкой (кг)": {"ozon": "Вес в упаковке, г*", "normalize": "kg_to_g"},
  "Высота упаковки": {"ozon": "Высота упаковки, мм*", "normalize": "cm_to_mm"},
  "Длина упаковки": {"ozon": "Длина упаковки, мм*", "normalize": "cm_to_mm"},
  "Ширина упаковки": {"ozon": "Ширина упаковки, мм*", "normalize": "cm_to_mm"},
  "Вид застежки": {"ozon": "Вид застежки"},
  "Тип посадки": {"ozon": "Тип посадки"},
  "Уход за вещами": {"ozon": "Уход за вещами"},
  "Ставка НДС": {"ozon": "НДС, %*"}
}
'''
mapping = json.loads(mapping_json)

# --- Функция поиска строки-заголовка по ключевому слову ---
# Удаляю функцию read_excel_with_header_search и все связанные с ней вызовы

def normalize_value(wb_row, wb_col, norm_rule, wb_df):
    # Защита от отсутствующих значений и нормализация
    value = wb_row.get(wb_col, "")
    if pd.isnull(value):
        return ""
    if norm_rule:
        if norm_rule.startswith('concat_with'):
            # Например: "concat_with|Рос. размер"
            _, with_col = norm_rule.split('|')
            with_value = wb_row.get(with_col, "")
            if pd.isnull(with_value): with_value = ""
            return f"{str(with_value).strip()} {str(value).strip()}".strip()
        elif norm_rule == 'kg_to_g':
            try:
                return round(float(value) * 1000, 2)
            except: return value
        elif norm_rule == 'cm_to_mm':
            try:
                return round(float(value) * 10, 2)
            except: return value
    return value

def update_ozon_file_only_data(original_path, updated_df):
    wb = openpyxl.load_workbook(original_path)
    ws = wb.worksheets[0]  # Первый лист
    # Сопоставляем названия столбцов с их индексами
    header_row = 1
    headers = [cell.value for cell in ws[header_row]]
    col_map = {name: idx for idx, name in enumerate(headers)}
    # Обновляем только те ячейки, которые изменились
    for i, row in updated_df.iterrows():
        for col_name, value in row.items():
            if col_name in col_map:
                col_idx = col_map[col_name] + 1  # openpyxl 1-based
                ws.cell(row=i + header_row + 1, column=col_idx, value=value)
    # Сохраняем в папку updated рядом с оригиналом
    updated_dir = os.path.join(os.path.dirname(original_path), 'updated')
    os.makedirs(updated_dir, exist_ok=True)
    save_path = os.path.join(updated_dir, os.path.basename(original_path))
    wb.save(save_path)
    return save_path

def process_files(wb_paths, ozon_paths, mapping):
    barcodes_aliases = mapping['Баркод'].get('aliases', [mapping['Баркод']['ozon'], 'Баркод'])
    for wb_path in wb_paths:
        # Чтение WB: заголовки на 3-й строке (header=2), 4-я строка — комментарии, удаляем её
        wb_df = pd.read_excel(wb_path, header=2)
        wb_df = wb_df.reset_index(drop=True)
        wb_df = wb_df.loc[:, ~wb_df.columns.str.contains('^Unnamed')]
        wb_df.columns = wb_df.columns.str.strip().str.lower()
        # Отладочный вывод столбцов
        print("Столбцы WB:", wb_df.columns.tolist())
        for col in wb_df.columns:
            print("Столбец:", repr(col))
        # Диагностика: вывод длины и кодов символов для всех столбцов с 'баркод'
        for col in wb_df.columns:
            if 'баркод' in col:
                print(f"Колонка: '{col}' | Длина: {len(col)} | Коды: {[ord(c) for c in col]}")
        # Очищаем названия столбцов от невидимых символов
        wb_df.columns = [re.sub(r'[^а-яa-z0-9 ]', '', col) for col in wb_df.columns]
        wb_df.columns = [col.strip().lower() for col in wb_df.columns]
        # Повторно ищем столбец для баркода
        wb_barcodes_col = next((col for col in wb_df.columns if 'баркод' in col), None)
        if wb_barcodes_col is None:
            for col in wb_df.columns:
                print(f"Столбец: '{col}' | Длина: {len(col)} | Коды: {[ord(c) for c in col]}")
            messagebox.showerror("Ошибка", f"В файле WB не найден столбец с баркодом даже после очистки. Найдены столбцы: {list(wb_df.columns)}")
            continue
        if wb_barcodes_col not in wb_df.columns:
            print("Столбец 'баркод' не найден!")
        else:
            print("Столбец 'баркод' найден, создаю копию DataFrame только с этим столбцом и нужными для работы.")
            cols = list(wb_df.columns)
            wb_df = wb_df[cols].copy()
            print("Столбцы после копирования:", wb_df.columns.tolist())
            # Финальный поиск столбца для индекса
            for col in wb_df.columns:
                print(f"Финальный столбец: '{col}'")
            final_barcodes_col = next((col for col in wb_df.columns if 'баркод' in col), None)
            print("Финальное имя столбца для индекса:", final_barcodes_col)
            wb_df = wb_df.set_index(final_barcodes_col)

        for ozon_path in ozon_paths:
            # Чтение Ozon: заголовки на 2-й строке (header=1)
            ozon_df = pd.read_excel(ozon_path, header=1)
            ozon_df = ozon_df.reset_index(drop=True)
            ozon_df.columns = ozon_df.columns.str.strip()
            ozon_barcodes_col = mapping['Баркод']['ozon']
            if ozon_barcodes_col not in ozon_df.columns:
                messagebox.showerror("Ошибка", f"В файле Ozon нет колонки '{ozon_barcodes_col}'\nНайдены столбцы: {list(ozon_df.columns)}")
                continue
            ozon_df = ozon_df.dropna(subset=[ozon_barcodes_col])
            ozon_df.set_index(ozon_barcodes_col, inplace=True)

            updated = 0
            for barcode, ozon_row in ozon_df.iterrows():
                if barcode in wb_df.index:
                    wb_row = wb_df.loc[barcode]
                    for wb_col, map_val in mapping.items():
                        ozon_col = map_val['ozon']
                        norm_rule = map_val.get('normalize')
                        # Проверяем, что нужные колонки существуют
                        if wb_col not in wb_df.columns or ozon_col not in ozon_df.columns:
                            continue
                        wb_value = normalize_value(wb_row, wb_col, norm_rule, wb_df)
                        ozon_value = ozon_row[ozon_col] if ozon_col in ozon_row else ""
                        # Если значение в Ozon не совпадает — меняем
                        if pd.notnull(wb_value) and str(ozon_value).strip() != str(wb_value).strip():
                            ozon_df.at[barcode, ozon_col] = wb_value
                            updated += 1

            save_path = update_ozon_file_only_data(ozon_path, ozon_df.reset_index())
            messagebox.showinfo('Результат',
                f"Обновлено {updated} полей в файле:\n{os.path.basename(save_path)}"
            )
            print(f"Сохранено: {save_path}")

def main():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo('WB -> Ozon нормализация', 'Выбери ВСЕ Excel-файлы WB')
    wb_paths = filedialog.askopenfilenames(title='Выбери Excel-файлы WB', filetypes=[('Excel files', '*.xlsx')])
    if not wb_paths:
        messagebox.showerror('Ошибка', 'Не выбраны файлы WB')
        return
    messagebox.showinfo('WB -> Ozon нормализация', 'Теперь выбери ВСЕ Excel-файлы Ozon')
    ozon_paths = filedialog.askopenfilenames(title='Выбери Excel-файлы Ozon', filetypes=[('Excel files', '*.xlsx')])
    if not ozon_paths:
        messagebox.showerror('Ошибка', 'Не выбраны файлы Ozon')
        return
    process_files(wb_paths, ozon_paths, mapping)
    messagebox.showinfo('Готово!', 'Обработка завершена. Проверь новые файлы с суффиксом "_updated".')

if __name__ == '__main__':
    main()
