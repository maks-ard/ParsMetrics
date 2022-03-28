import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec


class Base:
    def __init__(self, day: str, month='03', year='2022'):
        self.__URL = f'https://metrika.yandex.ru/stat/conversion_rate?no_robots=1&robots_metric=1&cross_device_attribution=1&cross_device_users_metric=0&group=dekaminute&period={year}-{month}-{day}%3A{year}-{month}-{day}&accuracy=1&id=19405381'

        options = Options()
        useragent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:97.0) Gecko/20100101 Firefox/97.0'
        options.add_argument(f'user-agent={useragent}')
        options.add_argument("--disable-client-side-phishing-detection")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-blink-features=AutomationControlled")
        # options.add_argument("--headless")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.__service = Service(executable_path=r'data/chromedriver.exe')
        self.driver = webdriver.Chrome(service=self.__service, options=options)
        self.__wait = WebDriverWait(self.driver, 5, 0.3)

    def autho_ym(self):
        self.driver.get(self.__URL)
        print("Авторизация")

        login = self.__wait.until(ec.visibility_of_element_located((By.XPATH, '//input[@type="text"]')))
        login.click()
        time.sleep(0.5)
        login.send_keys('ardeev.max@yandex.ru')
        login.send_keys(Keys.ENTER)

        password = self.__wait.until(ec.visibility_of_element_located((By.XPATH, '//input[@type="password"]')))
        time.sleep(1)
        password.click()
        password.send_keys('Thebesthimik2021')
        password.send_keys(Keys.ENTER)
        print('Авторизация успешна')

    def parse(self, day, month='03', year='2022'):
        ALL = f'https://metrika.yandex.ru/stat/traffic?goal=32946132&metric=ym%3As%3Agoal%3Cgoal_id%3Evisits&period={year}-{month}-{day}%3A{year}-{month}-{day}&accuracy=1&id=19405381&stateHash=6188d6af74684c00fe6de222'
        USERS = f'https://metrika.yandex.ru/stat/traffic?period={year}-{month}-{day}%3A{year}-{month}-{day}&accuracy=1&id=19405381&stateHash=6188e736f4778500b9cb836a'
        htmlfile = fr'data/source_pages/date_{day}.html'
        htmlfile_all_u = fr'data/source_pages/all_{day}.html'
        htmlfile_con_u = fr'data/source_pages/conv_{day}.html'

        self.driver.get(self.__URL)
        time.sleep(5)
        code_ym = self.driver.page_source
        with open(htmlfile, 'w', encoding='utf-8') as file:
            file.write(code_ym)

        self.driver.get(ALL)
        time.sleep(3)
        code_au = self.driver.page_source
        with open(htmlfile_all_u, 'w', encoding='utf-8') as file:
            file.write(code_au)

        self.driver.get(USERS)
        time.sleep(3)
        code_cu = self.driver.page_source
        with open(htmlfile_con_u, 'w', encoding='utf-8') as file:
            file.write(code_cu)

        self.driver.quit()
        print('Парсинг закончен...')
