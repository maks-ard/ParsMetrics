import json
import time
import openpyxl
from bs4 import BeautifulSoup
from progress_bar import print_progress_bar

filename = r'data/выгрузка.xlsx'
with open(r'data/ids.json', 'r', encoding='windows-1251') as file:
    all_id = json.load(file)
    len_dict = len(all_id)


def save_errors(errors):
    with open(r'data/errors.json', 'a', encoding='windows-1251') as er:
        json.dump(errors, er, indent=4, ensure_ascii=False)
        er.write('\n')


def read_index(day: str):
    metrics = {}
    errors = {day: {}}

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
    except Exception as ex:
        metrics['3'] = ''
        errors[str(day)]["3"] = {
            "id": all_id["3"]["id"],
            "name": all_id["3"]["name"],
            "message": f"Значение не найдено\n{ex}"
        }
        save_errors()

    # Поиск по ключу 4
    with open(htmlfile_con_u, encoding='utf-8') as file:
        src_4 = file.read()
    soup_4 = BeautifulSoup(src_4, 'lxml')

    try:
        element_4 = \
            soup_4.find_all('td', class_='data-table__cell data-table__cell_body_yes data-table__cell_type_metric')[4]
        num = ''.join(i for i in element_4.text if i in '0123456789')
        metrics['4'] = num
    except Exception as ex:
        metrics['4'] = ''
        errors[str(day)]["4"] = {
            "id": all_id["4"]["id"],
            "name": all_id["4"]["name"],
            "message": f"Значение не найдено\n{ex}"
        }

    # Поиск остальных ключей
    with open(htmlfile, encoding='utf-8') as file:
        src1 = file.read()
    soup = BeautifulSoup(src1, 'lxml')
    count = 0
    for key, dataid in all_id.items():
        time.sleep(0.01)
        print_progress_bar(int(key), len_dict)
        count += 1

        try:
            if key == '3' or key == '4':
                continue
            element = soup.find('tr', {'data-id': f'{dataid["id"]}'}).find(
                class_='conversion-report__goal-metric-row_type_visits').find('td',
                                                                              class_='conversion-report__goal-metric-row-right')
            num = ''.join(i for i in element.text if i in '0123456789')
            metrics[key] = num
        except Exception as ex:
            metrics[key] = ''
            errors[day][key] = {
                "id": all_id[str(key)]["id"],
                "name": all_id[str(key)]["name"],
                "message": f"Значение не найдено\n{ex}"
            }
    save_errors(errors)
    return metrics


def edit_file(day: str):
    metrics = read_index(day)
    book = openpyxl.load_workbook(filename=filename)
    sheet = book['Март']

    with open(r'data/date_col.json') as col:
        all_col = json.load(col)

    col = all_col[day]
    count = 0

    for key, index in metrics.items():
        count += 1
        if index != '':
            sheet[col + key] = int(index)
    book.save(filename=filename)
