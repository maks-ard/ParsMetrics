"""Редактирует файл с id, если что то меняется в таблице"""
import json

import openpyxl

from config import writing_ecxel


class Editor:
    def __init__(self):
        self.path = r'data/ids — копия.json'
        self.filename = writing_ecxel.file_for_write()
        self.book = openpyxl.load_workbook(self.filename)
        self.sheet = self.book[self.book.sheetnames[-2]]

    def get_all_name_column(self, col: str) -> list:
        """Возвращает полный список значение заданного столбца
        :param col: str буква стобца"""
        return [(cell.row, cell.value) for cell in self.sheet[col] if cell.value is not None]

    def get_list_names(self) -> list:
        """Возвращает список целей, где нужно заполнять значения"""
        return [i for i in self.get_all_name_column("A") if i[1][-1] != ':']

    def edit_rows_data(self):
        """Редактирует json файл с row целей на актуальные"""
        ignore_list = ["Кол-во заявок  ", "Посетители (Количество уникальных посетителей)"]
        update_row = [{"name": name, "row": row} for row, name in self.get_list_names()
                      if name not in ignore_list]

        with open(r"data/name-row.json", "w", encoding="utf-8") as file:
            json.dump(update_row, file, indent=4, ensure_ascii=False)

    def edit_ids_data(self):
        """Редактирует json файл с id целей на актуальные"""
        ids = json.load(open(r"data/ids.json", encoding="utf-8"))
        new_ids = {item["name"]: item["id"] for item in ids}

        with open(r"data/name-id.json", "w", encoding="utf-8") as file:
            json.dump(new_ids, file, indent=4, ensure_ascii=False)

    def get_name_id(self):
        ids = json.load(open(r"data/ids.json", encoding="utf-8"))
        list_names = []
        for i in ids:
            for key, item in i.items():
                list_names.append({
                    "name": key,
                    "row": item["row"],
                    "id": item["id"]
                })

        with open(r"data/ids.json", "w", encoding="utf-8") as file:
            json.dump(list_names, file, indent=4, ensure_ascii=False)

    def main(self):
        self.edit_rows_data()
