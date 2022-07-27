"""Редактирует файл с id, если что то меняется в таблице"""
import openpyxl


class GetIdRow:
    def __init__(self, filename):
        self.path = r'data/ids — копия.json'
        self.filename = filename
        self.book = openpyxl.load_workbook(self.filename)
        self.sheet = self.book["Июль"]

    def get_all_name_column(self, col: str) -> list:
        """Возвращает полный список значение заданного столбца
        :param col: str буква стобца"""
        return [(cell.row, cell.value) for cell in self.sheet[col] if cell.value is not None]

    def get_list_names(self) -> list:
        """Возвращает список целей, где нужно заполнять значения"""
        return [i for i in self.get_all_name_column("A") if i[1][-1] != ':']

    def get_id_row(self):
        return {item[1]: item[0] for item in self.get_all_name_column("BR")}
