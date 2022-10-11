import json
import openpyxl

from common.editor_excel import EditorExcel
from common import BaseExcel, api_yandex_async


class GeneralMetrics(BaseExcel):
    def __init__(self, filename=None):
        super(GeneralMetrics, self).__init__()
        self.filename = self.get_filepath("Метрики2022КОПИЯ") if filename is None else filename
        self.book = openpyxl.load_workbook(self.filename)

    def name_sheet(self, month=None):
        if month is None:
            month = self.get_yesterday()[1]
        return json.load(open(r"data/name_sheet.json", encoding="utf-8"))[month]  # текущий месяц

    def main(self, day=None, month=None, date1='yesterday', date2='yesterday'):
        if day is None:
            day = self.get_yesterday()[0]

        if month is None:
            month = self.get_yesterday()[1]

        ids: dict = EditorExcel(self.filename).get_id_row("BR")
        metrics = api_yandex_async.main(ids, date1=date1, date2=date2)  # выгруженные метрики

        sheet = self.book[self.name_sheet(month)]  # нужный лист в ecxel
        col = json.load(open(r'data/date_col.json', encoding="utf-8"))[day]  # нужная колонка

        for key, index in metrics.items():
            if key != "date" and index != '':
                sheet[col + str(key)] = int(index)  # запись значения в ячейки
        self.book.save(filename=self.filename)  # сохранение изменений
