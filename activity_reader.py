"""Модуль для чтения и обработки данных об активностях из Excel-файла."""
from parser import Parser


class ActivityReader:
    def __init__(self, file_path):
        self.file_path = file_path
        self.parser = Parser()
        self._data = None 

    def __get_data(self):
        if self._data is None:
            self._data = self.parser.parse_excel(self.file_path)
        return self._data

    def __get_unique(self):
        data = self.__get_data()

        seen = set()
        unique = []

        for row in data:
            activity = str(row[0]).strip()
            if (
                activity
                and "Итого" not in activity
                and "Проект" not in activity
                and "None" not in activity
                and "тариф" not in activity.lower()
                and activity not in seen
            ):
                seen.add(activity)
                unique.append(activity)

        return unique

    def read_activities(self):
        data = self.__get_data()
        activities = self.__get_unique()

        result = []

        for activity in activities:
            rows = [row for row in data if str(row[0]).strip() == activity]
            result.append((activity, rows))

        return result