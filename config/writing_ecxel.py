import json
import openpyxl
from config import api_yandex
from datetime import datetime

filename = r'data/выгрузка.xlsx'
yesterday = str(datetime.now().day - 1)


def edit_file(day=yesterday, date1='yesterday', date2='yesterday'):
    metrics = api_yandex.main(date1=date1, date2=date2)
    book = openpyxl.load_workbook(filename=filename)
    sheet = book['Март']

    with open(r'data/date_col.json') as col:
        all_col = json.load(col)

    col = all_col[day]

    for key, index in metrics.items():
        if index != '':
            sheet[col + key] = int(index)
    book.save(filename=filename)
