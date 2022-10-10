from common.editor_excel import EditorExcel
from config import api_yandex_async
from common import BaseExcel

class GeneralMetrics(BaseExcel):

    def main(self, day=get_yesterday("day"), month=get_yesterday("month"), date1='yesterday', date2='yesterday'):
        filename = file_for_write()  # путь к ecxel файлу
        name_sheet_now = name_sheet(month)
        ids: dict = EditorExcel(filename).get_id_row("BR")

        metrics = api_yandex_async.main(ids, date1=date1, date2=date2)  # выгруженные метрики
        book = openpyxl.load_workbook(filename=filename)  # книга ecxel

        sheet = book[name_sheet_now]  # нужный лист в ecxel
        col = json.load(open(r'data/date_col.json', encoding="utf-8"))[day]  # нужная колонка

        for key, index in metrics.items():
            if key != "date" and index != '':
                sheet[col + str(key)] = int(index)  # запись значения в ячейки
        book.save(filename=filename)  # сохранение изменений
