import openpyxl
from bs4 import BeautifulSoup
from selenium_ym import Base
from progress_bar import print_progress_bar

filename = r'data/выгрузка.xlsx'


all_id = {'3': '', '4': '',
          '6': '224297277', '7': '224639723', '8': '224834768', '9': '150874963', '11': '55411018', '12': '55411021',
          '13': '114998911',
          '14': '114998914', '16': '102628633',
          '17': '102628636', '18': '102628642', '19': '102628645', '21': '102677101', '22': '102677104',
          '23': '102677107', '24': '102677110', '26': '199430452',
          '27': '199430455', '28': '199430458', '29': '196198924', '31': '167643286', '32': '167643289',
          '33': '167643292', '34': '168294223', '35': '101217211',
          '37': '158922628', '39': '101322853', '40': '101322856', '41': '101322859', '43': '101328364',
          '44': '101328367', '45': '101328370', '46': '146128744',
          '47': '119690419', '49': '128708563', '50': '128708566', '52': '188721370', '53': '228873339',
          '54': '156370684', '57': '227367599', '58': '227367600',
          '60': '227368982', '61': '227368983', '62': '227368984', '63': '227368985', '64': '227368986',
          '65': '227392707', '66': '227392776', '68': '231551918', '69': '231551919', '71': '231791794',
          '72': '231791795', '73': '231791796', '74': '231791797', '75': '231791798'}


def read_index(day: str):
    metrics = {}
    htmlfile = fr'data/source_pages/date_{day}.html'
    htmlfile_all_u = fr'data/source_pages/all_{day}.html'
    htmlfile_con_u = fr'data/source_pages/conv_{day}.html'

    # Поиск по ключу 3
    with open(htmlfile_all_u, encoding='utf-8') as file:
        src_3 = file.read()
    soup_3 = BeautifulSoup(src_3, 'lxml')

    try:
        element_3 = \
            soup_3.find_all('td', class_='data-table__cell data-table__cell_body_yes data-table__cell_type_metric')[1]
        num = ''.join(i for i in element_3.text if i in '0123456789')
        metrics['3'] = num
    except Exception:
        metrics['3'] = ''
        print("'3' значение не найдено")

    # Поиск по ключу 4
    with open(htmlfile_con_u, encoding='utf-8') as file:
        src_4 = file.read()
    soup_4 = BeautifulSoup(src_4, 'lxml')

    try:
        element_4 = \
            soup_4.find_all('td', class_='data-table__cell data-table__cell_body_yes data-table__cell_type_metric')[4]
        num = ''.join(i for i in element_4.text if i in '0123456789')
        metrics['4'] = num
    except Exception:
        metrics['4'] = ''
        print("'4' значение не найдено")

    # Поиск остальных ключей
    with open(htmlfile, encoding='utf-8') as file:
        src1 = file.read()
    soup = BeautifulSoup(src1, 'lxml')

    for key, dataid in all_id.items():
        try:
            if key == '3' or key == '4':
                continue
            element = soup.find('tr', {'data-id': f'{dataid}'}).find(
                class_='conversion-report__goal-metric-row_type_visits').find('td',
                                                                              class_='conversion-report__goal-metric-row-right')
            num = ''.join(i for i in element.text if i in '0123456789')
            metrics[key] = num
            # print_progress_bar(int(key), 67)
        except Exception:
            metrics[key] = ''
            print(f'{key} значение не найдено...')
    return metrics


def edit_file(day: str):
    book = openpyxl.load_workbook(filename=filename)
    sheet = book.index()
    metrics = read_index()
    len_of = len(metrics)

    all_col = {
        '01': 'B', '02': 'D', '03': 'F', '04': 'H', '05': 'J', '06': 'L', '07': 'N', '08': 'P', '09': 'R', '10': 'T',
        '11': 'V', '12': 'X', '13': 'Z', '14': 'AB', '15': 'AD', '16': 'AF', '17': 'AH', '18': 'AJ', '19': 'AL',
        '20': 'AN',
        '21': 'AP', '22': 'AR', '23': 'AT', '24': 'AV', '25': 'AX', '26': 'AZ', '27': 'BB', '28': 'BD', '29': 'BF',
        '30': 'BH', '31': 'BJ'
    }

    # now_day = datetime.datetime.today().strftime("%d")
    # col = all_col[(str(int(now_day) - 1))]
    col = all_col[day]

    for key, index in metrics.items():
        if index != '':
            sheet[col + key] = int(index)
            print_progress_bar(int(key), len_of)

    book.save(filename=filename)


if __name__ == '__main__':
    base = Base('01')
    base.autho_ym()
    for date in range(18, 21):
        print(f'Закрузка файла -------> {date}')

        if 9 >= date >= 1:
            item = f'0{date}'

        base.parse(str(date))

    for item in range(18, 21):
        print(f'Запись файла -------> {item}')

        if 9 >= item >= 1:
            item = f'0{item}'

        edit_file(str(item))

    print(f'{"-" * 10}Done!{"-" * 10}')
