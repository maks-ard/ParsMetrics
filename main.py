import logging
import time
import traceback

import pandas as pd

from config import writing_ecxel
from config.editor import GetIdRow

logger = logging.getLogger("main")
logger.setLevel(logging.INFO)

file_handler = logging.FileHandler(r"data/pars_metrics.log", mode="w")

formatter = logging.Formatter("%(asctime)s : [%(levelname)s] [%(lineno)d] : %(message)s")
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)


def get_params():
    year = "2022"
    need_date = input(f"!!!ВСЕ ДАТЫ ПРОПИСЫВАЮТСЯ В ФОРМАТЕ DD MM YYYY!!!\n"
                      f"Год опционально, по-умолчанию стоит {year}\n"
                      f"Напиши первую дату если нужна выгрузка за период, иначе нажми Enter: ")
    if need_date == "":
        writing_ecxel.edit_file()

    elif need_date == "update":
        editor = GetIdRow(writing_ecxel.file_for_write())
        editor.update_date(writing_ecxel.name_sheet())
        editor.update_formulas(writing_ecxel.name_sheet())

    elif need_date == "refin":
        writing_ecxel.write_goal_refinance()

    else:

        start_date = need_date.split(" ")
        stop_date = input("Дата окончания d-m-y: ").split(" ")
        if len(start_date) == 3:
            year = start_date[2]
        daterange = pd.date_range(f'{year}-{start_date[1]}-{start_date[0]}', f'{year}-{stop_date[1]}-{stop_date[0]}')

        for date in daterange:
            date_now = str(date.strftime("%Y-%m-%d"))
            writing_ecxel.edit_file(day=date.strftime("%d"), month=date.strftime('%m'), date1=date_now, date2=date_now)


if __name__ == '__main__':
    try:
        start_time = time.time()
        get_params()
        # writing_ecxel.start_file()
        finish_time = time.time() - start_time
        print(f"TIME: {finish_time}")
    except Exception:
        print("Ошибка!")
        logger.critical(traceback.format_exc())
