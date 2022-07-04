import json
import platform

import openpyxl
from config import api_yandex_async

from datetime import datetime, timedelta
from tkinter import filedialog, Tk


def get_yesterday(day_or_month):
    return (datetime.now() - timedelta(days=1)).strftime('%d') if day_or_month == "day" else (
            datetime.now() - timedelta(days=1)).strftime('%m')


def start_file():  # открытие ecxel файла
    import os
    os.startfile(file_for_write())


def get_path_ecxel():  # открывает окно для выбора файла ecxel и возвращает полный путь
    root = Tk()
    root.attributes("-topmost", True)
    root.lift()
    root.withdraw()
    return filedialog.askopenfilename()


def choice_file(path_to_files):
    print("Выбери в какой файл записывать данные!")
    name_pc = platform.node()
    path = get_path_ecxel()
    path_to_files[name_pc] = path
    with open(r"data/path_to_ecxel.json", "w", encoding="utf-8") as file:
        json.dump(path_to_files, file, indent=3, ensure_ascii=False)
    return path


def file_for_write():  # выбор пути, в зависимости от устройства
    path_to_files = json.load(open(r"data/path_to_ecxel.json", encoding="utf-8"))

    try:
        file = open(path_to_files[platform.node()])
        file.close()
        return path_to_files[platform.node()]
    except KeyError:
        return choice_file()
    except FileNotFoundError:
        return choice_file(path_to_files)
    except PermissionError:
        import psutil

        for proc in psutil.process_iter():
            if proc.name() == "EXCEL.EXE":
                proc.kill()
        return path_to_files[platform.node()]


def edit_file(day=get_yesterday("day"), month=get_yesterday("month"), date1='yesterday', date2='yesterday'):
    filename = file_for_write()  # путь к ecxel файлу
    metrics = api_yandex_async.main(date1=date1, date2=date2)  # выгруженные метрики
    book = openpyxl.load_workbook(filename=filename)  # книга ecxel

    sheet = book[json.load(open(r"data/name_sheet.json", encoding="utf-8"))[month]]  # нужный лист в ecxel
    col = json.load(open(r'data/date_col.json', encoding="utf-8"))[day]  # нужная колонка

    for key, index in metrics.items():
        if index != '':
            sheet[col + key] = int(index)  # запись значения в ячейки
    book.save(filename=filename)  # сохранение изменений
