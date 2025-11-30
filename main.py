import argparse
from activity_reader import ActivityReader
from workbook_creator import WorkbookCreator


def main(file_path:str, output_dir:str):
    reader = ActivityReader(file_path)
    creator = WorkbookCreator()

    activities = reader.read_activities()

    for activity_name, rows in activities:
        if not rows:
            continue

        output_path = creator.create_workbook(rows, activity_name, output_dir=output_dir)
        print(f"Создан файл: {output_path}")


if __name__ == "__main__":
    argparser = argparse.ArgumentParser(description="Автоматизация создания внутренних отчетов из Excel-файлов.")
    argparser.add_argument("--file_path", type=str, help="Путь к исходному Excel-файлу.")
    argparser.add_argument("--output_dir", type=str, help="Директория для сохранения созданных отчетов.")
    args = argparser.parse_args()
    
    main(file_path=args.file_path, output_dir=args.output_dir)

