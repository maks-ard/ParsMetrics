import json
import openpyxl
from config import api_yandex
from datetime import datetime, timedelta

filename = r'C:\Users\Центрофинанс\OneDrive - ООО Микрокредитная компания «Центрофинанс Групп»\Рабочий стол\OneDrive - ООО Микрокредитная компания «Центрофинанс Групп»\Документы\Метрики2022КОПИЯ.xlsx'
yesterday = datetime.now() - timedelta(days=1)
all_month = {'01': 'Январь', '02': 'Февраль', '03': 'Март', '04': 'Апрель', '05': 'Май', '06': 'Июнь',
             '07': 'Июль', '08': 'Август', '09': 'Сентябрь', '10': 'Октябрь', '11': 'Ноябрь', '12': 'Декабрь'}


def edit_file(day=yesterday.strftime('%d'), month=yesterday.strftime('%m'), date1='yesterday', date2='yesterday'):
    metrics = api_yandex.main(date1=date1, date2=date2)
    book = openpyxl.load_workbook(filename=filename)
    sheet = book[all_month[month]]

    with open(r'data/date_col.json') as col:
        all_col = json.load(col)

    col = all_col[day]

    for key, index in metrics.items():
        if index != '':
            sheet[col + key] = int(index)
    book.save(filename=filename)
