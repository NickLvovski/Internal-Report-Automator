"""Модуль для создания новых выгрузок Excel."""
import calendar
from openpyxl import Workbook
from datetime import datetime
from pathlib import Path


class WorkbookCreator:
    def __init__(self):
        pass

    def create_workbook(self, data, activity_name, output_dir):
        total:float = 0
        output_dir = Path(output_dir) 
        output_dir.mkdir(parents=True, exist_ok=True)

        workbook = Workbook()
        sheet = workbook.active

        headers = [
            'Проект', 'Категория', 'Наименование услуги',
            'Тариф', 'Расчетный тариф', 'Количество',
            'Единица измерения', 'Стоимость'
        ]
        sheet.append(headers)

        for row in data:
            sheet.append(row)
            total += float(row[-1])

        sheet.append(["Итого", total])

        month_name = calendar.month_name[datetime.now().month]

        filename = f"Отчет_{activity_name}_{month_name}.xlsx"
        file_path = output_dir / filename

        workbook.save(file_path)
        return str(file_path)

