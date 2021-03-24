import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from employee_arrival_report import employee_arrival_report, title

work_directory = os.getcwd()

# Создаём файл для записи результата
wb = Workbook()
result_sheet = wb.active

# Устанавливаем заголовки столбцов
result_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
title_cell = result_sheet.cell(row=1, column=1)
title_cell.value = title
title_cell.font = Font(bold=True)

COLUMNS_NAMES = ['Фамилия', 'Имя', 'Отчество', 'Должность', 'Приход']
for index, column_name in enumerate(COLUMNS_NAMES):
    cell = result_sheet.cell(row=2, column=index+1)
    cell.value = column_name
    cell.font = Font(bold=True)
    cell.fill = PatternFill(bgColor='FFFF00', fill_type='gray0625')

# Устанавливаем ширину столбцов
COLUMNS = [('A', 12), ('B', 10), ('C', 14), ('D', 20), ('E', 10)]
for column, width in COLUMNS:
    result_sheet.column_dimensions[column].width = width

# Записываем данные таблицы
for row, employee in enumerate(employee_arrival_report()):
    for column, value in enumerate(employee):
        cell = result_sheet.cell(row=row+3, column=column+1)
        cell.value = value
        if employee[4] > '9:00:00':
            cell.fill = PatternFill(bgColor='FF6D6D', fill_type='gray0625')

wb.save(work_directory + '/result.xlsx')
