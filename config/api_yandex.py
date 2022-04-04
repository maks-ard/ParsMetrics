"""
ID: 1268a5e054fd4fbdbc4336a92cbbeb79
Пароль: 2f3f1516fb3a40dc9365f7f798c11ae6
Callback URL: https://oauth.yandex.ru/verification_code
Время жизни токена: Не менее, чем 1 год
Дата создания: 29.03.2022
"""
import json
import requests
from progress.progress.bar import IncrementalBar

TOKEN = "AQAAAABaprzMAAfIkr1IL8KgfEqppqmoi30MEUs"
headers = {'Authorization': f'OAuth {TOKEN}'}
metrics = {}

with open('data/ids.json', 'r', encoding='utf-8') as ids_file:
    ids = json.load(ids_file)


def get_users(date1='yesterday', date2='yesterday', id_counter=19405381):
    params = {
        'metrics': f'ym:s:users',
        'ids': id_counter,
        'date1': date1,
        'date2': date2
    }

    response = requests.get('https://api-metrika.yandex.net/stat/v1/data', headers=headers, params=params)
    users = response.json()
    all_users = int(users["totals"][0])
    return int(all_users)


def get_visits(goals, date1='yesterday', date2='yesterday', id_counter=19405381):
    params = {
        'metrics': f'{goals}',
        'ids': id_counter,
        'date1': date1,
        'date2': date2,
        "dimensions": "ym:s:date",
        'sort': "ym:s:date"
    }

    response = requests.get('https://api-metrika.yandex.net/stat/v1/data', headers=headers, params=params)
    users = response.json()
    if int(users["totals"][0]) == 0:
        return ''
    return int(users["totals"][0])


def main(date1='yesterday', date2='yesterday', id_counter=19405381):
    metrics["3"] = get_users(date1=date1, date2=date2, id_counter=id_counter)
    bar = IncrementalBar('Парсинг данных', max=len(ids))
    for key, value in ids.items():
        metrics[key] = get_visits(f'ym:s:goal{value["id"]}visits', date1=date1, date2=date2, id_counter=id_counter)
        bar.next()
    bar.finish()
    return metrics

