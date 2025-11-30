"""Модуль для чтения и обработки данных об активностях из Excel-файла."""
from parser import Parser


class ActivityReader:
    def __init__(self, file_path):
        self.file_path = file_path
        self.parser = Parser()
        self._data = None 

    def _get_data(self):
        if self._data is None:
            self._data = self.parser.parse_excel(self.file_path)
        return self._data

    def _get_unique(self):
        data = self._get_data()

        seen = set()
        unique = []

        for row in data:
            activity = str(row[0]).strip()
            if (
                activity
                and "Итого" not in activity
                and activity not in seen
            ):
                seen.add(activity)
                unique.append(activity)

        return unique

    def read_activities(self):
        data = self._get_data()
        activities = self._get_unique()

        result = []

        for activity in activities:
            rows = [row for row in data if str(row[0]).strip() == activity]
            result.append((activity, rows))

        return result