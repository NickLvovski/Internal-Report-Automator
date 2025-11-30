from activity_reader import ActivityReader
from workbook_creator import WorkbookCreator


def main():
    reader = ActivityReader("/Users/nfilatov/Python projects/Internal Report Automator/drpo.test.xlsx")
    creator = WorkbookCreator()

    activities = reader.read_activities()

    for activity_name, rows in activities:
        if not rows:
            continue

        file_path = creator.create_workbook(rows, activity_name)
        print(f"Создан файл: {file_path}")


if __name__ == "__main__":
    main()