"""Модуль для парсинга Excel файлов."""
import openpyxl


class Parser:
    def __init__(self):
        pass

    def parse_excel(self, file_path):
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        data = []

        for row in sheet.iter_rows(values_only=True):
            data.append(row)

        return data