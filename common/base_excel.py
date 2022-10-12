import json
import os
import platform
import logging
import traceback
import psutil
from datetime import datetime, timedelta
from tkinter import filedialog, Tk


class BaseExcel:

    def __init__(self):
        self.logger = logging.getLogger("main")
        self.filepath = r"data/path_to_excelfiles.json"

    @property
    def get_yesterday(self) -> datetime:
        return datetime.now() - timedelta(days=1)

    def start_file(self, filename):
        os.startfile(self.get_filepath(filename))

    def get_filepath(self, filename: str) -> str:
        """Получить путь до нужного файла"""
        all_files = json.load(open(self.filepath, "r", encoding="utf-8"))
        try:
            return all_files[filename]

        except (KeyError, FileNotFoundError):
            root = Tk()
            root.attributes("-topmost", True)
            root.lift()
            root.withdraw()

            all_files[filename] = filedialog.askopenfilename()

            json.dump(all_files, open(self.filepath, "w", encoding="utf-8"), indent=4, ensure_ascii=False)

            return filedialog.askopenfilename()

        except PermissionError:
            print("Файл для записи открыт!")

        except Exception:
            self.logger.critical(traceback.format_exc())
