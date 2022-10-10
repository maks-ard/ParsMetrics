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
        self.__filepath = r"data/path_to_ecxel.json"

    @staticmethod
    def get_yesterday() -> list[str]:
        """Возвращает вчерашнюю дату списком [dd, mm, yyyy]"""
        return (datetime.now() - timedelta(days=1)).strftime('%d.%m.%Y').split(".")

    def start_file(self, filename):
        """Открыть файл"""
        os.startfile(self.get_filepath(filename))

    def get_filepath(self, filename: str) -> str:
        """Получить путь до нужного файла"""
        all_files = json.load(open(self.__filepath, encoding="utf-8"))
        try:
            return all_files[filename]

        except (KeyError, FileNotFoundError):
            root = Tk()
            root.attributes("-topmost", True)
            root.lift()
            root.withdraw()

            all_files[filename] = filedialog.askopenfilename()

            json.dump(all_files, open(self.__filepath, encoding="utf-8"), indent=4, ensure_ascii=False)

            return filedialog.askopenfilename()

        except PermissionError:
            print("Файл для записи открыт!")

        except Exception:
            self.logger.critical(traceback.format_exc())


