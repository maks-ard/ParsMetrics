"""
Callback URL: https://oauth.yandex.ru/verification_code
Время жизни токена: Не менее, чем 1 год
Дата создания: 29.03.2022
"""
import asyncio
import json
from types import SimpleNamespace

import aiohttp
import requests
import logging

from common.privat_info import TOKEN

logger = logging.getLogger("main")
metrics = {}


class YandexApi:
    def __init__(self, id_counter=19405381):
        self.base_url = "https://api-metrika.yandex.net/stat/v1/"
        self.id_counter = id_counter

    @staticmethod
    def get_response_object(data):
        return json.loads(json.dumps(data.json()), object_hook=lambda d: SimpleNamespace(**d))

    @property
    def headers(self):
        return {'Authorization': f'OAuth {TOKEN}'}

    def get_params(self, metric, first_date, last_date, **kwargs):
        return {
            'metrics': metric,
            'ids': self.id_counter,
            'date1': first_date,
            'date2': last_date
        } | kwargs

    def get_total_time(self, date):
        url = self.base_url + "data"
        params = self.get_params("ym:s:avgVisitDurationSeconds", date, date)

        response = self.get_response_object(requests.get(url, headers=self.headers, params=params))
        return round(int(response.totals[0]) / 60, 2)

    def get_users(self, date1='yesterday', date2='yesterday'):
        url = self.base_url + "data"
        params = self.get_params(f'ym:s:users', date1, date2)
        response = self.get_response_object(requests.get(url, headers=self.headers, params=params))

        return response.totals[0]

    async def get_visits(self, session, row, goals, date1='yesterday', date2='yesterday'):
        url = self.base_url + "data"

        params = self.get_params(f'ym:s:goal{goals}visits', date1, date2)
        async with session.get(url, headers=self.headers, params=params) as response:
            users = await response.json()
            if 200 <= response.status <= 399:
                metrics[row] = (users["totals"][0])
                metrics["date"] = users["query"]["date1"]

            elif response.status == 400:
                logger.warning(users)
                logger.info(params)

            else:
                response.raise_for_status()

    def get_csat(self, date):
        url = self.base_url + "data"

        params = self.get_params("ym:s:paramsNumber",
                                 date,
                                 date,
                                 dimensions="ym:s:paramsLevel2",
                                 sort="ym:s:paramsLevel2",
                                 filters="ym:s:paramsLevel1=='ratingVote'")

        response = requests.get(url, headers=self.headers, params=params)

        data = response.json()["data"]

        return {item["dimensions"][0]["name"]: item["metrics"][0] for item in data}

    async def gather_data(self, ids: dict, date1='yesterday', date2='yesterday'):
        async with aiohttp.ClientSession(connector=aiohttp.TCPConnector(limit=3)) as session:
            tasks = list()
            for id_goal, row in ids.items():
                tasks.append(
                    asyncio.create_task(
                        self.get_visits(session, row, id_goal, date1=date1, date2=date2)
                    )
                )
            await asyncio.gather(*tasks)


def main(ids: dict, date1='yesterday', date2='yesterday', need_users=True):
    global metrics
    metrics = {}
    api = YandexApi()

    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(api.gather_data(ids, date1=date1, date2=date2))

    if need_users:
        metrics["3"] = api.get_users(date1=date1, date2=date2)

    return metrics
