"""
ID: 1268a5e054fd4fbdbc4336a92cbbeb79
Пароль: 2f3f1516fb3a40dc9365f7f798c11ae6
Callback URL: https://oauth.yandex.ru/verification_code
Время жизни токена: Не менее, чем 1 год
Дата создания: 29.03.2022
"""

import json

import requests


def save_json(obj):
    with open('responce1.json', 'w', encoding='utf-8') as file:
        json.dump(obj, file, indent=4, ensure_ascii=False)


def save_html(obj):
    with open('responce.html', 'w', encoding='utf-8') as file:
        file.write(obj)


TOKEN = "AQAAAABaprzMAAfIkr1IL8KgfEqppqmoi30MEUs"

headers = {'Authorization': f'OAuth {TOKEN}'}


def get_users():
    params = {
        'metrics': f'ym:s:users',
        'ids': 19405381,
        'date1': 'yesterday',
        'date2': 'yesterday'
    }

    response = requests.get('https://api-metrika.yandex.net/stat/v1/data', headers=headers, params=params)
    users = response.json()
    all_users = int(users["totals"][0])
    print(all_users)


def get_visits(id_goal):
    params = {
        'metrics': f'ym:s:goa{id_goal}users',
        'ids': 19405381,
        'date1': 'yesterday',
        'date2': 'yesterday'
    }

    response = requests.get('https://api-metrika.yandex.net/stat/v1/data', headers=headers, params=params)
    users = response.json()
    # all_users = int(users["totals"][0])
    print(users)


def get_counters():
    results = {}
    response = requests.get('https://api-metrika.yandex.net/management/v1/counter/19405381/goals', headers=headers)
    goals = response.json()
    save_json(goals)
    for goal in goals["goals"]:
        if goal["is_retargeting"] == 1:
            results[goal["name"]] = goal["id"]
    print(results)


def main():
    get_visits(32946132)


if __name__ == '__main__':
    main()
