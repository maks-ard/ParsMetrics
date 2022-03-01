import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec


def parse_metrics(day: str):
    URL = f'https://metrika.yandex.ru/stat/conversion_rate?no_robots=1&robots_metric=1&cross_device_attribution=1&cross_device_users_metric=0&period=2022-02-{day}%3A2022-02-{day}&accuracy=1&id=19405381'
    URL_CONVERS = f'https://metrika.yandex.ru/stat/traffic?goal=32946132&metric=ym%3As%3Agoal%3Cgoal_id%3Evisits&period=2022-02-{day}%3A2022-02-{day}&accuracy=1&id=19405381&stateHash=6188d6af74684c00fe6de222'
    URL_USERS = f'https://metrika.yandex.ru/stat/traffic?period=2022-02-{day}%3A2022-02-{day}&accuracy=1&id=19405381&stateHash=6188e736f4778500b9cb836a'

    options = Options()
    useragent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:97.0) Gecko/20100101 Firefox/97.0'
    options.add_argument(f'user-agent={useragent}')
    # options.add_argument("--headless")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    service = Service(executable_path=r'data/chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 5, 0.3)

    try:
        print('Парсер запущен...')
        driver.get(URL)

        login = wait.until(ec.visibility_of_element_located((By.XPATH, '//input[@type="text"]')))
        login.click()
        time.sleep(0.5)
        print("Ввод логина...")
        login.send_keys('ardeev.max@yandex.ru')
        login.send_keys(Keys.ENTER)

        password = wait.until(ec.visibility_of_element_located((By.XPATH, '//input[@type="password"]')))
        time.sleep(1)
        password.click()
        print("Ввод пароля...")
        password.send_keys('Thebesthimik2021')
        password.send_keys(Keys.ENTER)
        time.sleep(5)

        code_ym = driver.page_source
        print('Запись index.html')
        with open(r'data/index.html', 'w', encoding='utf-8') as file:
            print('Запись страницы в HTML файл...')
            file.write(code_ym)

        driver.get(URL_USERS)
        time.sleep(3)
        code_au = driver.page_source
        print('Запись index_all.html')
        with open(r'data/index_all.html', 'w', encoding='utf-8') as file:
            print('Запись страницы в HTML файл...')
            file.write(code_au)

        driver.get(URL_CONVERS)
        time.sleep(3)
        code_cu = driver.page_source
        print('Запись index_conversed.html')
        with open(r'data/index_conversed.html', 'w', encoding='utf-8') as file:
            print('Запись страницы в HTML файл...')
            file.write(code_cu)

        driver.quit()
        print('Парсинг данных успешно завершён!')
    except Exception as error:
        print(error)
        driver.quit()
        parse_metrics()
