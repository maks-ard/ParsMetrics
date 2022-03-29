"""
ID: 1268a5e054fd4fbdbc4336a92cbbeb79
Пароль: 2f3f1516fb3a40dc9365f7f798c11ae6
Callback URL: https://oauth.yandex.ru/verification_code
Время жизни токена: Не менее, чем 1 год
Дата создания: 29.03.2022
"""

import requests
import json

TOKEN = "AQAAAABaprzMAAfIkr1IL8KgfEqppqmoi30MEUs"

headers = {'Authorization': f'OAuth {TOKEN}'}


params = {
    "is_retargeting": 1
}

response = requests.get('https://api-metrika.yandex.net/management/v1/counter/19405381/goals', headers=headers, params=params)
goals = response.json()

for goal in goals["goals"]:
    if goal["is_retargeting"] == 1:
        print(goal["name"])
# with open('responce.json', 'w', encoding='utf-8') as file:
#     json.dump(goals, file, indent=4, ensure_ascii=False)
