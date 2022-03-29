from config import main, selenium_ym
import os

if os.path.isfile('data/errors.json'):
    os.remove('data/errors.json')
start_day = int(input('Введите дату старта: '))
end_day = input('Введите дату окончания (включительно): ')
base = selenium_ym.Base(str(start_day))

try:
    if end_day == '':
        base.autho_ym()
        base.parse(start_day, month='03')
        main.edit_file(str(start_day))
        base.exit()
    else:
        base.autho_ym()
        for date in range(start_day, int(end_day) + 1):
            print(f'Закрузка файла -------> {date}')

            if 9 >= date >= 1:
                item = f'0{date}'

            base.parse(str(date))
        base.exit()
        for item in range(start_day, int(end_day) + 1):
            print(f'Запись файла -------> {item}')

            if 9 >= item >= 1:
                item = f'0{item}'

            main.edit_file(str(item))
        print(f'{"-" * 10}Done!{"-" * 10}')
except Exception as ex:
    print(ex)
    base.exit()
