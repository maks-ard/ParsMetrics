"""
Callback URL: https://oauth.yandex.ru/verification_code
Время жизни токена: Не менее, чем 1 год
Дата создания: 29.03.2022
"""
import asyncio
import aiohttp
import requests
import logging

from progress.bar import IncrementalBar
from config.privat_info import TOKEN

logger = logging.getLogger("main")
URL = "https://api-metrika.yandex.net/stat/v1/data"
headers = {'Authorization': f'OAuth {TOKEN}'}
metrics = {}
count = 0

global bar


def get_users(date1='yesterday', date2='yesterday', id_counter=19405381):
    params = {
        'metrics': f'ym:s:users',
        'ids': id_counter,
        'date1': date1,
        'date2': date2
    }

    response = requests.get(URL, headers=headers, params=params)
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
    async with session.get(URL, headers=headers, params=params) as response:
        users = await response.json()
        if 200 <= response.status <= 399:
            metrics[row] = (users["totals"][0])
            metrics["date"] = users["query"]["date1"]
            bar.next()

        elif response.status == 400:
            logger.warning(users)
            logger.info(params)
            bar.next()

        else:
            response.raise_for_status()


async def gather_data(ids: dict, date1='yesterday', date2='yesterday', id_counter=19405381):
    connector = aiohttp.TCPConnector(limit=3)  # Ограничивает количество параллельных запросов
    async with aiohttp.ClientSession(connector=connector) as session:
        tasks = []
        for id_goal, row in ids.items():
            task = asyncio.create_task(
                get_visits(session, row, f'ym:s:goal{id_goal}visits', date1=date1, date2=date2,
                           id_counter=id_counter))
            tasks.append(task)
        await asyncio.gather(*tasks)


def main(ids: dict, date1='yesterday', date2='yesterday', id_counter=19405381, need_users=True):
    global bar
    bar = IncrementalBar(f'Парсинг данных: {date1}', max=len(ids))

    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(gather_data(ids, date1=date1, date2=date2, id_counter=id_counter))
    if need_users:
        metrics["3"] = get_users(date1=date1, date2=date2, id_counter=id_counter)

    bar.finish()
    return metrics
