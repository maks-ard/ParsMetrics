import openpyxl, datetime
from bs4 import BeautifulSoup
import lxml
import selenium_ym

filename = 'metrics.xlsx'
all_id = {'6': '224297277', '7': '224639723', '8': '224834768', '11': '55411018', '12': '55411021', '13': '114998911',
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
          '65': '227392707', '66': '227392776'}


def read_index():
    metrics = {}
    with open(r'index.html', encoding='utf-8') as file:
        src = file.read()
    soup = BeautifulSoup(src, 'lxml')
    for key, dataid in all_id.items():
        try:
            element = soup.find('tr', {'data-id': f'{dataid}'}).find(
                class_='conversion-report__goal-metric-row_type_visits').find('td', class_='conversion-report__goal-metric-row-right')
            num = ''.join(i for i in element.text if i in '0123456789')
            metrics[key] = num
        except Exception:
            title = soup.find('tr', {'data-id': f'{dataid}'}).find(class_='conversion-report__goal-title')
            metrics[key] = ''
            print(f'Ошибка в {title.text}')
    return metrics


def edit_file():
    book = openpyxl.load_workbook(filename=filename)
    sheet = book.active
    metrics = read_index()

    all_col = {
        '01': 'B', '02': 'D', '03': 'F', '04': 'H', '05': 'J', '06': 'L', '07': 'N', '08': 'P', '09': 'R', '10': 'T',
        '11': 'V', '12': 'X', '13': 'Z', '14': 'AB', '15': 'AD', '16': 'AF', '17': 'AH', '18': 'AJ', '19': 'AL',
        '20': 'AN',
        '21': 'AP', '22': 'AR', '23': 'AT', '24': 'AV', '25': 'AX', '26': 'AZ', '27': 'BB', '28': 'BD', '29': 'BF',
        '30': 'BH', '31': 'BJ'
    }

    now_day = datetime.datetime.today().strftime("%d")
    col = all_col[now_day]

    for key, index in metrics.items():
        sheet[col + key] = index

    book.save(filename=filename)


if __name__ == '__main__':
    selenium_ym.parse_metrics()
    edit_file()
