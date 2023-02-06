import logging
import sys
import time
import traceback
from datetime import datetime
import plyer

import pandas as pd
import schedule

from lib import GeneralMetrics, RefinancingExcel
from common.editor_excel import EditorExcel

general = GeneralMetrics()
editor = EditorExcel(general.filename)

service_name = "parser-yandex-metrics"

def auto():
    refinancing = RefinancingExcel()
    refinancing.main()
    schedule.every(1).hour.do(refinancing.main)
    # schedule.every(1).days.do(general.main)


def get_logger():
    logger = logging.getLogger(service_name)
    logger.setLevel(logging.INFO)

    stream_handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return logger

def get_params():
    year = "2022"

    choice = input(f"Что выполнить?: ")

    if choice == "general":
        general.main()

    elif choice == "refin":
        refinancing = RefinancingExcel()
        refinancing.main()

    elif choice == "update":
        editor.update_date(general.name_sheet())
        editor.update_formulas(general.name_sheet())

    elif choice == "auto":
        # general.main()
        # refinancing.main()
        print(datetime.now())
        auto()
        while True:
            schedule.run_pending()
            time.sleep(60)

    else:
        start_date = choice.split(" ")
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
            general.start_file(general.filename)

        finish_time = time.time() - start_time
        logger.info(f"TIME: {finish_time}")

    except Exception:
        plyer.notification.notify(message='Error download metrics!',
                                  app_name='ParsMetrics',
                                  app_icon='data/static/problem_icon-icons.com_54943.ico',
                                  title='Error',
                                  timeout=5)
        logger.critical(traceback.format_exc())


if __name__ == '__main__':
    logger = get_logger()
    main(startfile=False)
