import json
import os
import platform
import logging
import traceback
import psutil
from datetime import datetime, timedelta
from tkinter import filedialog, Tk


class BaseExcel:

    def __int__(self):
        self.logger = logging.getLogger("main")

    def get_yesterday(self, day_or_month):
        return (datetime.now() - timedelta(days=1)).strftime('%d') if day_or_month == "day" else (
                datetime.now() - timedelta(days=1)).strftime('%m')

    def get_path_ecxel(self):  # открывает окно для выбора файла ecxel и возвращает полный путь
        root = Tk()
        root.attributes("-topmost", True)
        root.lift()
        root.withdraw()

        return filedialog.askopenfilename()

    def start_file(self, filename=None):  # открытие ecxel файла
        if filename is None:
            os.startfile(self.file_for_write())
        else:
            os.startfile(filename)



    def choice_file(self, path_to_files):
        print("Выбери в какой файл записывать данные!")
        name_pc = platform.node()
        path = self.get_path_ecxel()
        path_to_files[name_pc] = path

        with open(r"data/path_to_ecxel.json", "w", encoding="utf-8") as file:
            json.dump(path_to_files, file, indent=3, ensure_ascii=False)

        return path

    def file_for_write(self):  # выбор пути, в зависимости от устройства
        path_to_files = json.load(open(r"data/path_to_ecxel.json", encoding="utf-8"))

        try:
            for proc in psutil.process_iter():
                if proc.name() == "EXCEL.EXE":
                    proc.kill()
            file = open(path_to_files[platform.node()])  # проверяет корректность пути до файла и не открыт ли он
            file.close()
            return path_to_files[platform.node()]

        except (KeyError, FileNotFoundError):
            self.logger.error(traceback.format_exc())
            return self.choice_file(path_to_files)

    def name_sheet(self, month=None):
        if month is None:
            month = datetime.now().strftime("%m")
        return json.load(open(r"data/name_sheet.json", encoding="utf-8"))[month]  # текущий месяц

    def get_row_for_write(self, sheet):
        rows = sheet.iter_rows(max_col=1)
        result = ()
        for row in rows:
            for cell in row:
                result = (cell.value, cell.row)
                if cell.value is None:
                    return cell.offset(row=-1).value, cell.offset(row=-1).row
        return result
