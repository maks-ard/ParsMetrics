"""Редактирует файл с id, если что то меняется в таблице"""
import json

import openpyxl

from config import writing_ecxel

IGNORE_GOAL = []

class Editor:
    def __init__(self):
        self.path = r'data/ids — копия.json'
        self.filename = writing_ecxel.file_for_write()
        self.book = openpyxl.load_workbook(self.filename)
        self.sheet = self.book[self.book.sheetnames[-2]]
        old_ids = json.load(open(self.path, encoding="utf-8"))
        new_ids = {}

    def get_all_name_column(self, col: str):
        return [cell.value for cell in self.sheet[col]]

    def get_ignore_list(self):
        names = self.get_all_name_column("A")
        values_b = self.get_all_name_column("B")

        for name in names:
            if values_b[names.index(name)] is None:
                IGNORE_GOAL.append(name)

        print(IGNORE_GOAL)

