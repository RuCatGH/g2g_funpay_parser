import time
import multiprocessing
import pickle

from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from openpyxl import Workbook

seller_names = []  # Список имен продавцов


wb = Workbook()
ws = wb.active

options = webdriver.ChromeOptions()
options.add_argument(f'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--window-size=1920,1080")
options.headless = True

ws.append(['Seller id', 'Game', 'Price', 'Date'])

def parse_feedbacks(url: str) -> None:
    
    
    driver = webdriver.Chrome(options=options)

    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        'source': """
            delete cdc_adoQpoasnfa76pfcZLmcfl_Array;
            delete cdc_adoQpoasnfa76pfcZLmcfl_Promise;
            delete cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
            delete cdc_adoQpoasnfa76pfcZLmcfl_Proxy;
            delete cdc_adoQpoasnfa76pfcZLmcfl_Object;
        """
    })



    driver.get(url)
    games = driver.find_elements(By.CLASS_NAME, 'list-inline')  # Найти все элементы игр
    all_links_games = []
    for game in games:
        links = [link.get_attribute('href') for link in game.find_elements(By.TAG_NAME, 'a')]  # Получить ссылки на игры
        all_links_games.extend(links)

    num_cores = multiprocessing.cpu_count() # Получение количества ядер процессора
    pool = multiprocessing.Pool(processes=num_cores)

    pool.map(get_all_offers, all_links_games)
    pool.close()
    pool.join()

    driver.quit()  # Закрытие драйвера после выполнения всех задач для данного ядра

def get_all_offers(link: str) -> None:
    try:
    
        driver = webdriver.Chrome(options=options)  # Создание экземпляра драйвера для каждого ядра

        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            'source': """
                delete cdc_adoQpoasnfa76pfcZLmcfl_Array;
                delete cdc_adoQpoasnfa76pfcZLmcfl_Promise;
                delete cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
                delete cdc_adoQpoasnfa76pfcZLmcfl_Proxy;
                delete cdc_adoQpoasnfa76pfcZLmcfl_Object;
            """
        })

        driver.get(link)
        # Загружаем cookies
        for cookies in pickle.load(open("cookies", "rb")):
            driver.add_cookie(cookies)

        driver.refresh()
        h1 = driver.find_element(By.TAG_NAME, 'h1').text == '429 Too Many Requests'
        while h1:
            time.sleep(1)
            driver.refresh()
            h1 = driver.find_element(By.TAG_NAME, 'h1').text == '429 Too Many Requests'

        for offer in driver.find_elements(By.CLASS_NAME, 'tc-item'):  # Найти все предложения
            seller_name = offer.find_element(By.CLASS_NAME, 'media-user-name').get_attribute('textContent')  # Имя продавца
            if seller_name not in seller_names:
                seller_names.append(seller_name)
                url = offer.get_attribute('href')  # Ссылка на предложение
                get_feedbacks(driver, url)
                wb.save('feedbacks.xlsx')

        
    except Exception as ex:
        print(ex)
    finally:
        driver.close()
        driver.quit()  # Закрытие драйвера после выполнения всех задач для данного ядра
def get_feedbacks(driver, link: str):
    driver.execute_script("window.open('" + link + "');")
    driver.switch_to.window(driver.window_handles[-1])

    h1 = driver.find_element(By.TAG_NAME, 'h1').text == '429 Too Many Requests'
    while h1:
        time.sleep(1)
        driver.refresh()
        h1 = driver.find_element(By.TAG_NAME, 'h1').text == '429 Too Many Requests'

    seller_id = driver.find_element(By.CLASS_NAME, 'chat').get_attribute('data-seller')  # ID продавца
    while True:
        limiter_date = get_info_feedbacks(driver, seller_id)
        if limiter_date:
            break
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

def get_info_feedbacks(driver, seller_id):
    comments = driver.find_elements(By.CLASS_NAME, 'review-item')  # Найти все отзывы
    for comment in comments:
        date = comment.find_element(By.CLASS_NAME, 'review-item-date').get_attribute('textContent').strip()  # Дата отзыва
        main_info = comment.find_element(By.CLASS_NAME, 'review-item-detail').get_attribute('textContent').strip()  # Основная информация
        price = main_info.split(',')[-1].strip()  # Цена
        game = ','.join(main_info.split(',')[:-1]).strip()  # Игра

        if date not in ['В этом месяце', 'Месяц назад']:
            return True

        data = [seller_id, game, price, date]  # Данные отзыва

        ws.append(data)  # Добавить данные в таблицу

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    try:
        driver.find_element(By.CLASS_NAME, 'dyn-table-continue').click()
    except NoSuchElementException:
        return True

if __name__ == "__main__":
    try:
        parse_feedbacks('https://funpay.com/')
    except KeyboardInterrupt:
        wb.save('feedbacks.xlsx')
