from config import writing_ecxel
import pandas as pd


def get_params():
    need_date = input("Нужна выгрузка за несколько дней?: ")
    if need_date == "":
        writing_ecxel.edit_file()
    else:
        start_date = input("Дата старта: ").split(" ")
        stop_date = input("Дата окончания: ").split(" ")
        daterange = pd.date_range(f'2022-{start_date[1]}-{start_date[0]}', f'2022-{stop_date[1]}-{stop_date[0]}')

        for date in daterange:
            date_now = str(date.strftime("%Y-%m-%d"))
            writing_ecxel.edit_file(day=date.strftime("%d"), month=date.strftime('%m'), date1=date_now, date2=date_now)


if __name__ == '__main__':
    get_params()
    writing_ecxel.start_file()
