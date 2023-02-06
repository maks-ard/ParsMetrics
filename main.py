import logging
import sys
import time
import traceback
from datetime import datetime
import plyer

from argparse import ArgumentParser
import schedule

from lib import RefinancingExcel

service_name = "parser-yandex-metrics"


def get_logger(level):
    log = logging.getLogger(service_name)
    log.setLevel(level * 10)

    stream_handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    stream_handler.setFormatter(formatter)
    log.addHandler(stream_handler)

    return log


def create_parser():
    parser = ArgumentParser()

    modes = ["auto", "refin"]
    current_year = str(datetime.today().year)

    parser.add_argument("-m", "--mode", default="auto", type=str, choices=modes)
    parser.add_argument("-s", "--startfile", default=False, const=True, action="store_const")
    parser.add_argument("-y", "--year", default=current_year, type=str)
    parser.add_argument("-l", "--loglevel", default=2, type=int)

    return parser


if __name__ == '__main__':
    args = create_parser().parse_args()
    logger = get_logger(args.loglevel)
    refinancing = RefinancingExcel()

    try:
        if args.mode == "refin":
            refinancing.main()

        elif args.mode == "auto":
            schedule.every(1).hour.do(refinancing.main)
            refinancing.main()
            while True:
                schedule.run_pending()
                time.sleep(1)

    except Exception:
        plyer.notification.notify(message='Error download metrics!',
                                  app_name='ParsMetrics',
                                  app_icon='data/static/problem_icon-icons.com_54943.ico',
                                  title='Error',
                                  timeout=5)
        logger.critical(traceback.format_exc())
