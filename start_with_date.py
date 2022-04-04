from config import writing_ecxel
import pandas as pd


def get_params():
    start_day = input('День старта dd: ')
    start_month = input('Месяц старта mm: ')
    start_year = '2022'
    stop_day = input('День окончания dd: ')
    stop_month = input('Месяц окончания mm: ')
    stop_year = '2022'
    daterange = pd.date_range(f'{start_year}-{start_month}-{start_day}', f'{stop_year}-{stop_month}-{stop_day}')

    for date in daterange:
        date_now = str(date.strftime("%Y-%m-%d"))
        writing_ecxel.edit_file(day=date.strftime("%d"), month=date.strftime('%m'), date1=date_now, date2=date_now)


if __name__ == '__main__':
    get_params()
