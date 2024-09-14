import os
import re
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from datetime import datetime

# Функция для применения стиля к ячейкам, включая границы
def apply_cell_style_with_borders(cell, bold=False, align_center=True, fill_color=None, border=None):
    # Настройка жирного шрифта, если указано
    cell.font = Font(bold=bold)
    # Выравнивание по центру, если нужно
    if align_center:
        cell.alignment = Alignment(horizontal="center", vertical="center")
    # Установка заливки, если указан цвет
    if fill_color:
        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    # Установка границ, если они переданы
    if border:
        cell.border = border

# Функция для форматирования листа Excel
def format_output_sheet(sheet, date_obj, border):
    # Установка ширины колонок
    sheet.column_dimensions['A'].width = 20  # Для колонки "№ тепловоза"
    sheet.column_dimensions['B'].width = 20  # Для колонки "Цех"
    sheet.column_dimensions['C'].width = 50  # Для колонки "Наименование работ"
    sheet.column_dimensions['D'].width = 10  # Для колонки "План"
    sheet.column_dimensions['E'].width = 10  # Для колонки "Факт"

    # Установка заголовка с датой
    sheet.merge_cells('A1:E1')
    sheet['A1'] = f"План-задание на {date_obj.strftime('%d.%m.%Y')}"
    apply_cell_style_with_borders(sheet['A1'], bold=True, align_center=True)
    sheet['A1'].font = Font(size=14)  # Увеличиваем шрифт заголовка

    # Установка заголовков колонок
    headers = ['№ тепловоза', 'Цех', 'Наименование работ', 'План', 'Факт']
    for col_num, header in enumerate(headers, 1):
        # Заполнение заголовков в строке 2
        cell = sheet.cell(row=2, column=col_num, value=header)
        apply_cell_style_with_borders(cell, bold=True, border=border)

# Функция для проверки значимых строк
def is_significant_row(row):
    # Проверка наличия данных в колонках '№ тепловоза' и 'Наименование'
    return pd.notna(row['№ тепловоза']) and pd.notna(row['Наименование']) and str(row['Наименование']).strip() != ""

# Функция для очистки текста от лишних символов
def clean_text(text):
    # Если текст отсутствует, возвращаем пустую строку
    if pd.isna(text):
        return ""
    # Удаление подчеркиваний, дефисов и пробелов
    return re.sub(r'[_-]+', '', str(text)).strip()

# Функция для проверки на пустую или нежелательную строку
def is_row_empty_or_unwanted(ws, row_index, num_columns):
    # Проходим по всем ячейкам строки
    for col in range(1, num_columns + 1):
        value = ws.cell(row=row_index, column=col).value
        # Если ячейка содержит значимые данные, возвращаем False
        if value not in [None, "", "_", "-", "ЦКТ", "ЦПМ", "МСЦ", "ЭМУ"]:
            return False
    # Если строка пустая или содержит нежелательные данные, возвращаем True
    return True

# Функция для получения даты с листа Excel
def get_sheet_date(sheet):
    # Чтение значения ячейки A1
    date_str = sheet['A1'].value
    if date_str:
        # Поиск даты в формате дд.мм.гггг
        date_match = re.search(r'\d{2}\.\d{2}\.\d{4}', date_str)
        if date_match:
            # Преобразование найденной строки в объект даты
            try:
                return datetime.strptime(date_match.group(), '%d.%m.%Y')
            except ValueError:
                return None
    # Если дата не найдена, возвращаем None
    return None

# Функция для сортировки листов по дате
def sort_worksheets_by_date(wb):
    date_sheets = {}
    # Перебор всех листов в книге
    for sheet in wb.sheetnames:
        try:
            # Преобразование имени листа в дату
            date_obj = datetime.strptime(sheet, '%d.%m')
            date_sheets[sheet] = date_obj
        except ValueError:
            pass  # Пропускаем листы без даты

    # Сортируем листы по дате
    sorted_sheets = sorted(date_sheets.items(), key=lambda x: x[1])

    # Меняем порядок листов в книге
    for i, (sheetname, _) in enumerate(sorted_sheets):
        sheet = wb[sheetname]
        wb.move_sheet(sheet, i + 1)

# Запрос даты у пользователя
date_str = input('Введите нужную дату в формате дд.мм.гггг: ')

try:
    # Преобразование строки в дату
    target_date = datetime.strptime(date_str, '%d.%m.%Y')
    print(f"Вы ввели дату: {target_date.strftime('%d.%m.%Y')}")
except ValueError:
    # Сообщение об ошибке, если дата введена неверно
    print("Неверный формат даты. Пожалуйста, используйте формат дд.мм.гггг.")
    exit()

# Путь к папкам с исходными файлами и выходной папке
inputDir = r'D:\Сменные задания+заявки ОТК'
outputDir = r'D:\Сменные задания+заявки ОТК\План-задание по цехам'

# Создаем папку для выходного файла, если она не существует
if not os.path.exists(outputDir):
    os.makedirs(outputDir)

# Определяем стиль границ
thin = Side(border_style="thin", color="000000")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# Названия заводов
factories = ['ЦКТ', 'ЦПМ', 'МСЦ', 'ЭМУ']
plan_filename = os.path.join(outputDir, "План-задание.xlsx")

# Если файл существует, загружаем его, иначе создаем новый
if os.path.exists(plan_filename):
    wb = load_workbook(plan_filename)
else:
    wb = Workbook()

# Удаляем лист "Sheet", если он пустой
if 'Sheet' in wb.sheetnames:
    sheet = wb['Sheet']
    if sheet.max_row == 1 and sheet.max_column == 1 and sheet.cell(1, 1).value is None:
        wb.remove(sheet)

# Создаем имя листа на основе даты
sheet_name = target_date.strftime('%d.%m')
if sheet_name in wb.sheetnames:
    # Удаляем лист, если он уже существует
    ws = wb[sheet_name]
    wb.remove(ws)

# Создаем новый лист
ws = wb.create_sheet(title=sheet_name)

# Форматируем лист
format_output_sheet(ws, target_date, border)

# Инициализируем индекс строки для записи данных
row_index = 3

# Словарь для хранения данных по тепловозам и заводам
loco_data = {}

# Обрабатываем файлы для каждого завода
for factory in factories:
    factory_dir = os.path.join(inputDir, factory)
    for root, dirs, files in os.walk(factory_dir):
        for filename in files:
            if filename.startswith('~$'):  # Пропускаем временные файлы
                continue

            if filename.endswith(".xlsx"):
                file_path = os.path.join(root, filename)

                try:
                    # Загружаем файл Excel
                    workbook = load_workbook(file_path, data_only=True)

                    for sheet_name in workbook.sheetnames:
                        sheet = workbook[sheet_name]
                        sheet_date = get_sheet_date(sheet)

                        # Если дата листа совпадает с введенной пользователем
                        if sheet_date and sheet_date == target_date:
                            df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)

                            # Проходим по строкам DataFrame
                            for i, r in df.iterrows():
                                if is_significant_row(r):
                                    loco_number = clean_text(r['№ тепловоза'])
                                    work_name = clean_text(r['Наименование'])

                                    if loco_number not in loco_data:
                                        loco_data[loco_number] = {}
                                    if factory not in loco_data[loco_number]:
                                        loco_data[loco_number][factory] = []
                                    loco_data[loco_number][factory].append(work_name)

                except Exception as e:
                    print(f"Ошибка при обработке файла {filename}: {e}")

# Записываем данные в итоговый файл
for loco_number, factory_data in loco_data.items():
    first_row = row_index

    for factory, works in factory_data.items():
        # Записываем название завода в колонку "Цех"
        ws.cell(row=row_index, column=2, value=factory)
        # Применяем стиль с границами к ячейке "Цех"
        apply_cell_style_with_borders(ws.cell(row=row_index, column=2), border=border)
        factory_first_row = row_index  # Сохраняем первую строку для завода
        # Проходим по всем работам, связанным с заводом
        for work_name in works:
            # Записываем название работы в колонку "Наименование работ"
            ws.cell(row=row_index, column=3, value=work_name)
            # Применяем стиль с границами к ячейке "Наименование работ"
            apply_cell_style_with_borders(ws.cell(row=row_index, column=3), border=border)
            row_index += 1  # Переход к следующей строке
            # Если для завода было несколько работ, объединяем ячейки для завода
        if factory_first_row < row_index - 1:
            ws.merge_cells(start_row=factory_first_row, start_column=2, end_row=row_index - 1, end_column=2)
            # Применяем стиль к объединенной ячейке
            apply_cell_style_with_borders(ws.cell(row=factory_first_row, column=2), border=border)
    # Объединяем ячейки для тепловоза, если заводов было несколько
    ws.merge_cells(start_row=first_row, start_column=1, end_row=row_index - 1, end_column=1)
    # Записываем номер тепловоза в колонку "№ тепловоза"
    ws.cell(row=first_row, column=1, value=loco_number)
    # Применяем стиль с границами к ячейке "№ тепловоза"
    apply_cell_style_with_borders(ws.cell(row=first_row, column=1), border=border)


# Добавляем границы для пустых столбцов "План" и "Факт"
for i in range(3, row_index):
    apply_cell_style_with_borders(ws.cell(row=i, column=4), border=border)  # Применяем границы к "План"
    apply_cell_style_with_borders(ws.cell(row=i, column=5), border=border)  # Применяем границы к "Факт"

# Удаляем пустые строки и строки с только нежелательными значениями
for i in range(row_index, 0, -1):
    if is_row_empty_or_unwanted(ws, i, 5):
        ws.delete_rows(i, 1)
    else:
        break

# Сортируем листы в книге по дате
sort_worksheets_by_date(wb)

# Сохраняем итоговый файл Excel
wb.save(plan_filename)
print(f"Обработанный файл сохранен в: {plan_filename}")
