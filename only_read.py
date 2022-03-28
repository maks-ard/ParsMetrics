from config import main
import os

if os.path.isfile('data/errors.json'):
    os.remove('data/errors.json')

start_day = input('Введите дату старта: ')
end_day = input('Введите дату окончания (включительно): ')

try:
    if end_day == '':
        main.edit_file(start_day)
    else:
        for item in range(int(start_day), int(end_day) + 1):
            print(f'Запись файла -------> {item}')

            if 9 >= item >= 1:
                item = f'0{item}'

            main.edit_file(str(item))
        main.save_errors()
        print(f'{"-" * 10}Done!{"-" * 10}')
except Exception as ex:
    print(ex)
