import json
import openpyxl
from config import api_yandex
from datetime import datetime, timedelta

yesterday = datetime.now() - timedelta(days=1)
all_month = {'01': 'Январь', '02': 'Февраль', '03': 'Март', '04': 'Апрель', '05': 'Май', '06': 'Июнь',
             '07': 'Июль', '08': 'Август', '09': 'Сентябрь', '10': 'Октябрь', '11': 'Ноябрь', '12': 'Декабрь'}


def file_for_write():
    import platform
    work_pc = r"C:\Users\m.ardeev\OneDrive - ООО Микрокредитная компания «Центрофинанс Групп»\Документы\Метрики2022КОПИЯ.xlsx"
    laptop = r"C:\Users\Центрофинанс\OneDrive - ООО Микрокредитная компания «Центрофинанс Групп»\Рабочий стол\OneDrive - ООО Микрокредитная компания «Центрофинанс Групп»\Документы\Метрики2022КОПИЯ.xlsx"
    if platform.node() == "M-ARDEEV-N":
        return laptop
    elif platform.node() == "WKS-IT-004":
        return work_pc
    else:
        return input("Введите путь до ecxel файла: ")


def start_file():
    import os
    os.startfile(file_for_write())


def edit_file(day=yesterday.strftime('%d'), month=yesterday.strftime('%m'), date1='yesterday', date2='yesterday'):
    filename = file_for_write()
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
