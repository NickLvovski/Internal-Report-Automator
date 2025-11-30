"""Модуль для создания новых выгрузок Excel."""
import openpyxl
from openpyxl import Workbook
from datetime import datetime
import calendar


class WorkbookCreator:
    def __init__(self):
        pass

    def create_workbook(self, data, activity_name):
        workbook = Workbook()
        sheet = workbook.active
        first_row = ['Проект', 'Категория', 'Наименование услуги',
                    'Тариф', 'Расчетный тариф', 'Количество', 'Единица измерения', 'Стоимость']
        sheet.append(first_row)
        
        for row in data:
            sheet.append(row)
        
        month_name = calendar.month_name[datetime.now().month]
        file_path = "Отчет_" + activity_name + f"_{month_name}.xlsx"
        workbook.save(file_path)
        return file_path

