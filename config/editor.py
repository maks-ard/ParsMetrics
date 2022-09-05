"""Редактирует файл с id, если что то меняется в таблице"""
"""В процессе"""
import json
import openpyxl
import pandas as pd
from datetime import datetime
from calendar import monthrange

from config import writing_ecxel


class GetIdRow:
    def __init__(self, filename):
        self.path = r'data/ids — копия.json'
        self.filename = filename
        self.book = openpyxl.load_workbook(self.filename)
        self.sheet = self.book["Август"]

    def get_all_name_column(self, col: str) -> list:
        """Возвращает полный список значение заданного столбца
        :param col: str буква стобца"""
        return [(cell.row, cell.value) for cell in self.sheet[col] if cell.value is not None]

    def get_list_names(self) -> list:
        """Возвращает список целей, где нужно заполнять значения"""
        return [i for i in self.get_all_name_column("A") if i[1][-1] != ':']

    def get_id_row(self):
        """Получает id целей"""
        return {item[1]: item[0] for item in self.get_all_name_column("BR")}


    def copy_list(self, title=None):
        """Копирует лист для нового месяца"""
        if title is None:
            name_sheet = json.load(open(r"data/name_sheet.json", encoding="utf-8"))
            month_now = datetime.now().strftime("%m")
            title = name_sheet[month_now]  # Название устанавливается согласно текущему месяцу

        new_sheet = self.book.copy_worksheet(self.sheet)
        new_sheet.title = title
        self.update_date(new_sheet)

        self.book.save(self.filename)

    def update_date(self, name_sheet: str = None, year: str = None, month: str = None):
        """Обновляет строки с датой на актуальные для листа
        :param name_sheet: лист, где надо обновить даты
        :param year: str год формата yyyy
        :param month: str месяц формата mm"""

        if name_sheet is None:
            name_sheet = writing_ecxel.name_sheet()

        if month is None:
            month = datetime.now().strftime("%m")
        if year is None:
            year = datetime.now().year

        sheet = self.book[name_sheet]

        days = monthrange(year, int(month))[1]

        month_limit = {
            29: 71,
            30: 73,
            31: 75
        }
        result = [chr(i) for i in range(66, 91, 2)] + \
                 ['A' + chr(i) for i in range(66, 91, 2)] + \
                 ['B' + chr(i) for i in range(66, month_limit[days], 2)]
        daterange = pd.date_range(f"{year}-{month}-01", f"{year}-{month}-{days}").strftime("%d.%m.20%y")

        day = 0
        for item in result:
            sheet[item + "1"] = daterange[day]
            sheet[item + "73"] = daterange[day]

            day += 1

        self.book.save(self.filename)

    def update_formulas(self, name_sheet: str = None, year: str = None, month: str = None):

        if name_sheet is None:
            name_sheet = writing_ecxel.name_sheet()

        if month is None:
            month = datetime.now().strftime("%m")
        if year is None:
            year = datetime.now().year

        sheet = self.book[name_sheet]
        days = monthrange(year, int(month))[1]

        work_columns = {
            29: ["BH", "BJ"],
            30: ["BJ", "BL"],
            31: ["BL", "BN"]
        }

        for column in work_columns[days]:
            for cell in sheet[column]:
                if cell.value is not None:
                    if cell.value[0] == "=":
                        result = ",".join(cell.value.split(",")[:-1]) + ")"
                        sheet[cell.coordinate] = result

        self.book.save(self.filename)

