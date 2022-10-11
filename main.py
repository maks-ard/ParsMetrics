import logging
import time
import traceback

import pandas as pd

from lib import GeneralMetrics, RefinancingExcel
from common.editor_excel import EditorExcel

general = GeneralMetrics()
refinancing = RefinancingExcel(filename=r"tests/CR Перекредитование в ЛК.xlsx")
editor = EditorExcel(general.get_filepath(general.filename))
editor = EditorExcel(general.get_filepath(general.filename))


def get_logger():
    log = logging.getLogger("main")
    log.setLevel(logging.INFO)

    file_handler = logging.FileHandler(r"data/pars_metrics.log", mode="w")

    formatter = logging.Formatter("%(asctime)s : [%(levelname)s] [%(lineno)d] : %(message)s")
    file_handler.setFormatter(formatter)

    log.addHandler(file_handler)


def get_params():
    year = "2022"
    need_date = input(f"!!!ВСЕ ДАТЫ ПРОПИСЫВАЮТСЯ В ФОРМАТЕ DD MM YYYY!!!\n"
                      f"Год опционально, по-умолчанию стоит {year}\n"
                      f"Напиши первую дату если нужна выгрузка за период, иначе нажми Enter: ")
    if need_date == "":
        general.main()

    elif need_date == "update":
        editor.update_date(general.name_sheet())
        editor.update_formulas(general.name_sheet())

    elif need_date == "refin":
        refinancing.main()

    else:
        start_date = need_date.split(" ")
        stop_date = input("Дата окончания d-m-y: ").split(" ")
        if len(start_date) == 3:
            year = start_date[2]
        daterange = pd.date_range(f'{year}-{start_date[1]}-{start_date[0]}', f'{year}-{stop_date[1]}-{stop_date[0]}')

        for date in daterange:
            date_now = str(date.strftime("%Y-%m-%d"))
            general.main(day=date.strftime("%d"), month=date.strftime('%m'), date1=date_now, date2=date_now)


def main(startfile=False):
    try:
        start_time = time.time()
        get_params()
        if startfile:
            general.start_file()
        finish_time = time.time() - start_time
        logger.info(f"TIME: {finish_time}")

    except Exception:
        print("Ошибка!")
        logger.critical(traceback.format_exc())


if __name__ == '__main__':
    logger = logging.getLogger("main")
    main(startfile=False)
