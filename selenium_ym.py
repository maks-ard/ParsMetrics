import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

url = 'https://metrika.yandex.ru/stat/conversion_rate?id=19405381&period=yesterday&accuracy=1'


def parse_metrics():
    options = Options()
    useragent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:97.0) Gecko/20100101 Firefox/97.0'
    options.add_argument(f'user-agent={useragent}')
    options.add_argument("--headless")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    service = Service(executable_path=r'data/chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 5, 0.3)

    try:
        driver.get(url)
        login = wait.until(ec.visibility_of_element_located((By.XPATH, '//input[@type="text"]')))
        login.click()
        time.sleep(0.5)
        login.send_keys('ardeev.max@yandex.ru')
        login.send_keys(Keys.ENTER)
        password = wait.until(ec.visibility_of_element_located((By.XPATH, '//input[@type="password"]')))
        time.sleep(1)
        password.click()
        password.send_keys('Thebesthimik2021')
        password.send_keys(Keys.ENTER)
        time.sleep(3)
        code_ym = driver.page_source
        with open(r'data/index.html', 'w', encoding='utf-8') as file:
            file.write(code_ym)
        driver.quit()
    except Exception as error:
        print(error)
        driver.quit()
        parse_metrics()
