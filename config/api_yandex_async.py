"""
ID: 1268a5e054fd4fbdbc4336a92cbbeb79
Пароль: 2f3f1516fb3a40dc9365f7f798c11ae6
Callback URL: https://oauth.yandex.ru/verification_code
Время жизни токена: Не менее, чем 1 год
Дата создания: 29.03.2022
"""
import asyncio
import aiohttp
import requests

from progress.bar import IncrementalBar

TOKEN = "AQAAAABaprzMAAfIkr1IL8KgfEqppqmoi30MEUs"
headers = {'Authorization': f'OAuth {TOKEN}'}
metrics = {}
count = 0


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


async def get_visits(session, row, goals, date1='yesterday', date2='yesterday', id_counter=19405381):
    params = {
        'metrics': f'{goals}',
        'ids': id_counter,
        'date1': date1,
        'date2': date2,
        "dimensions": "ym:s:date",
        'sort': "ym:s:date"
    }
    global count
    count += 1
    if count == 2:
        await asyncio.sleep(1)
        count = 0
    async with session.get('https://api-metrika.yandex.net/stat/v1/data', headers=headers, params=params) as response:
        users = await response.json()
        try:
            if int(users["totals"][0]) == 0:
                metrics[row] = ''
            metrics[row] = (users["totals"][0])
            bar.next()
        except KeyError:
            await get_visits(session, row, goals, date1='yesterday', date2='yesterday', id_counter=19405381)


async def gather_data(ids, date1='yesterday', date2='yesterday', id_counter=19405381):
    async with aiohttp.ClientSession() as session:
        tasks = []
        for id_goal, row in ids.items():
            task = asyncio.create_task(
                get_visits(session, row, f'ym:s:goal{id_goal}visits', date1=date1, date2=date2,
                           id_counter=id_counter))
            tasks.append(task)
        await asyncio.gather(*tasks)


def main(ids, date1='yesterday', date2='yesterday', id_counter=19405381):
    global bar
    bar = IncrementalBar(f'Парсинг данных: {date1}', max=len(ids))

    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(gather_data(ids, date1=date1, date2=date2, id_counter=id_counter))

    metrics["3"] = get_users(date1=date1, date2=date2, id_counter=id_counter)

    bar.finish()
    return metrics
