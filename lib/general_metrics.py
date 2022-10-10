import json
import openpyxl

from common.editor_excel import EditorExcel
from common import BaseExcel, api_yandex_async


class GeneralMetrics(BaseExcel):
    def __int__(self):
        self.filename = "Метрики2022КОПИЯ"

    def name_sheet(self, month=None):
        if month is None:
            month = self.get_yesterday()[1]
        return json.load(open(r"data/name_sheet.json", encoding="utf-8"))[month]  # текущий месяц

    def main(self, date1='yesterday', date2='yesterday'):
        filename = self.get_filepath(self.filename)
        ids: dict = EditorExcel(filename).get_id_row("BR")

        metrics = api_yandex_async.main(ids, date1=date1, date2=date2)  # выгруженные метрики
        book = openpyxl.load_workbook(filename=filename)  # книга ecxel

        sheet = book[self.name_sheet(self.get_yesterday()[1])]  # нужный лист в ecxel

        col = json.load(open(r'data/date_col.json', encoding="utf-8"))[self.get_yesterday()[0]]  # нужная колонка

        for key, index in metrics.items():
            if key != "date" and index != '':
                sheet[col + str(key)] = int(index)  # запись значения в ячейки
        book.save(filename=filename)  # сохранение изменений
