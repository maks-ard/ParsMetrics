import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

url = 'https://metrika.yandex.ru/stat/conversion_rate?id=19405381&period=yesterday&accuracy=1'


def parse_metrics():
    options = Options()
    useragent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:97.0) Gecko/20100101 Firefox/97.0'
    options.add_argument(f'user-agent={useragent}')
    service = Service(executable_path=r'chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_window_position(0, 0)
    driver.set_window_size(2000, 1000)
    try:
        driver.get(url)
        time.sleep(1)
        login = driver.find_element(By.XPATH, '//input[@type="text"]')
        login.click()
        time.sleep(1)
        login.send_keys('ardeev.max@yandex.ru')
        time.sleep(1)
        login.send_keys(Keys.ENTER)
        time.sleep(2)
        password = driver.find_element(By.XPATH, '//input[@type="password"]')
        time.sleep(1)
        password.click()
        password.send_keys('Thebesthimik2021')
        time.sleep(1)
        password.send_keys(Keys.ENTER)
        time.sleep(5)
        code_ym = driver.page_source
        with open('index.html', 'w', encoding='utf-8') as file:
            file.write(code_ym)
        driver.quit()
    except Exception as error:
        print(error)
        driver.quit()
