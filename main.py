import logging
import sys
import time
import traceback
from datetime import datetime

from argparse import ArgumentParser
import schedule
from dotenv import load_dotenv

from lib import RefinancingExcel

load_dotenv()
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

    current_year = str(datetime.today().year)

    parser.add_argument("-s", "--startfile", default=False, const=True, action="store_const")
    parser.add_argument("-o", "--only_one", default=False, const=True, action="store_const")
    parser.add_argument("-y", "--year", default=current_year, type=str)
    parser.add_argument("-l", "--loglevel", default=2, type=int)

    return parser


if __name__ == '__main__':
    args = create_parser().parse_args()
    logger = get_logger(args.loglevel)
    refinancing = RefinancingExcel()

    if args.only_one:
        try:
            logger.info("Единоразовый запуск")
            refinancing.main()

        except KeyboardInterrupt:
            logger.warning("Ручная остановка")
        except Exception:
            logger.critical(traceback.format_exc())
    else:
        try:
            schedule.every().day.at("09:30").do(refinancing.main)
            schedule.every().day.at("14:00").do(refinancing.main)
            schedule.every().day.at("16:30").do(refinancing.main)

            if args.startfile:
                refinancing.main()

            while True:
                schedule.run_pending()
                time.sleep(1)

        except KeyboardInterrupt:
            logger.warning("Ручная остановка")

        except Exception:
            logger.critical(traceback.format_exc())
