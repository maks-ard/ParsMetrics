"""Редактирует файл с id, если что то меняется в таблице"""
import openpyxl
import pandas as pd


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
        return {item[1]: item[0] for item in self.get_all_name_column("BR")}

    def update_date(self, month):
        month_limit = {
            29: 71,
            30: 73,
            31: 75
        }

        result = [chr(i) for i in range(66, 91, 2)] + ['B' + chr(i) for i in range(66, month_limit[month], 2)]
        daterange = pd.date_range("2022-08-01", "2022-08-31").strftime("%d.%m.20%y")

        day = 0
        for item in result:
            self.sheet[item + "1"] = daterange[day]
            self.sheet[item + "67"] = daterange[day]

            day += 1
        self.book.save(self.filename)
