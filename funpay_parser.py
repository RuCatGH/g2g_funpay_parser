import time
import os
import pickle
import json

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from dotenv import load_dotenv
from openpyxl import Workbook
import requests
import jmespath

from utils import headers, titles_for_xlsx

class FunPayParser:
    def __init__(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument(f'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36')
        self.options.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.options.add_experimental_option('useAutomationExtension', False)
        self.options.add_argument("--disable-blink-features=AutomationControlled")
        self.options.add_argument("--window-size=1920,1080")
    
        self.driver = webdriver.Chrome(options=self.options)

        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            'source': """
                delete cdc_adoQpoasnfa76pfcZLmcfl_Array;
                delete cdc_adoQpoasnfa76pfcZLmcfl_Promise;
                delete cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
                delete cdc_adoQpoasnfa76pfcZLmcfl_Proxy;
                delete cdc_adoQpoasnfa76pfcZLmcfl_Object;
            """
        })

        self.wait = WebDriverWait(self.driver, 5)


        load_dotenv()
        self.login = os.getenv('login')
        self.password = os.getenv('password')

        self.inputs_price = ['1000', '10000']

        self.wb = Workbook()
        
        self.ws = self.wb.active
        self.ws.append(titles_for_xlsx)

        self.s = requests.Session()


    def get_cookies_authorization(self, driver: webdriver.Chrome) -> dict:
        """
        Получает cookies после авторизации и сохраняет их в файле.
        Возвращает словарь с cookies.
        """
        driver.get('https://funpay.com/account/login')
        button = driver.find_element(By.CSS_SELECTOR, '.btn.btn-primary.btn-block')
        while button:
            time.sleep(1)
            try:
                button = driver.find_element(By.CSS_SELECTOR, '.btn.btn-primary.btn-block')
            except NoSuchElementException:
                button = None
            
        pickle.dump(driver.get_cookies(), open("cookies", "wb"))
        return driver.get_cookies()

    def parse(self, url: str) -> None:
        """
        Основной метод для парсинга данных с FunPay.
        Параметр url - URL страницы для парсинга.
        """
        try:
            # Получение cookies
            if not os.path.exists('cookies'):
                self.get_cookies_authorization(self.driver)

            self.driver.get(url)

            selenium_cookies = pickle.load(open("cookies", "rb"))
            # Загружаем cookies для сессии
            for cookie in selenium_cookies:
                self.s.cookies.set(cookie['name'], cookie['value'])

            # Загружаем cookies
            for cookies in pickle.load(open("cookies", "rb")):
                self.driver.add_cookie(cookies)

            self.driver.refresh()
            time.sleep(3)

            items = self.driver.find_elements(By.CLASS_NAME, 'promo-game-item')
            for item in items:
                title = item.find_element(By.CSS_SELECTOR, ".game-title:not([class*=' '])").find_element(By.TAG_NAME, 'a').text
                breadcrumbs1 = item.find_elements(By.CLASS_NAME, 'btn')

                # Выполнение скроллинга к элементу
                self.move_to_element(item)
                
                # Если есть секция 1 с выбором региона
                if breadcrumbs1:
                    for button in breadcrumbs1:
                        button.click()
                        links = item.find_element(By.CSS_SELECTOR, ".list-inline:not([class*=' '])").find_elements(By.TAG_NAME, 'li')
                        for link in links:
                            category_name = link.find_element(By.TAG_NAME, 'a').text
                            data = [title] + [button.text] + [category_name] + self.data_retrieval(self.driver, link)
                            self.save_to_excel(data)
                # Если нет первой секции с выбором региона
                else:
                    links = item.find_element(By.CSS_SELECTOR, ".list-inline:not([class*=' '])").find_elements(By.TAG_NAME, 'li')
                    for link in links:
                        category_name = link.find_element(By.TAG_NAME, 'a').text
                        data = [title] + ['None'] + [category_name] + self.data_retrieval(self.driver, link)
                        self.save_to_excel(data)
        except Exception  as ex:
            print(ex)
        finally:
            self.driver.close()
            self.driver.quit()
    
    def move_to_element(self, element) -> None:
        """
        Делает скролл до элемента так, чтобы он был по середине экрана
        """
        actions = ActionChains(self.driver)
        actions.move_to_element(element).perform()

        # Получение размеров окна браузера
        window_size = self.driver.get_window_size()
        window_width = window_size['width']
        window_height = window_size['height']

        # Получение позиции элемента на странице
        element_position = element.location_once_scrolled_into_view
        element_x = element_position['x']
        element_y = element_position['y']

        # Рассчет смещения для прокрутки элемента в центр экрана
        offset_x = int(window_width / 2 - element_x)
        offset_y = int(window_height / 2 - element_y)

        # Выполнение прокрутки элемента в центр экрана
        self.driver.execute_script(f"window.scrollBy({offset_x}, {offset_y});")


    def get_filters(self, driver: webdriver.Chrome) -> list[dict]:
        """
        Получает фильтры из страницы и возвращает список фильтров.
        Параметр driver - экземпляр WebDriver.
        Возвращает список словарей с данными фильтров.
        """
        filters_data: list[dict] = []

        # Получение элемента с фильтрами
        filters = driver.find_element(By.CLASS_NAME, 'showcase-filters')

        # Перебор каждого фильтра
        for filter in filters.find_elements(By.CSS_SELECTOR, ".form-group:not([class*=' '])"):

            try:
                switch_filter = filter.find_element(By.CSS_SELECTOR, '.form-control-box.switch')
            except NoSuchElementException:
                switch_filter = None

            if switch_filter:
                # Обработка фильтра-переключателя
                text_switch_filter = switch_filter.find_element(By.TAG_NAME, 'span').text
                filters_data.append({text_switch_filter: ['False', 'True']})
            else:
                # Обработка остальных типов фильтров
                sub_filters = [sub_filter.text for sub_filter in filter.find_elements(By.TAG_NAME, 'option')]
                filters_data.append({sub_filters[0]: sub_filters[1:]})

        try:
            json_fields = json.loads(filters.find_element(By.CSS_SELECTOR, '.lot-fields.live').get_attribute('data-fields'))
        except NoSuchElementException:
            json_fields = None

        # Перебор фильтров из data-fields
        if json_fields:
            id_fields = [element["id"] for element in json_fields]
            for id_field in id_fields:
                sub_filter = filters.find_element(By.XPATH, f'//*[@data-id="{id_field}"]')
                
                # Фильтр от и до
                try:
                    range_filter = sub_filter.find_elements(By.CLASS_NAME, 'lot-field-range-box')
                except NoSuchElementException:
                    range_filter = None

                # Фильтр список
                try:
                    choose_filter = sub_filter.find_elements(By.TAG_NAME, 'option')
                except NoSuchElementException:
                    choose_filter = None
                # Фильтр кнопка
                try:
                    button_filter = sub_filter.find_elements(By.TAG_NAME, 'button')
                except NoSuchElementException:
                    button_filter = None

                # Обработка фильтра взависимости от типа
                if range_filter:
                    range_filter_text = sub_filter.find_element(By.CLASS_NAME, 'control-label').get_attribute('textContent')
                    filters_data.append({range_filter_text: ['Min', 'Max']})
                elif choose_filter:
                    sub_choose_filters = [sub_choose_filter.text for sub_choose_filter in choose_filter]
                    filters_data.append({sub_choose_filters[0]: sub_choose_filters[1:]})
                elif button_filter:
                    sub_button_filters = [sub_button_filter.text for sub_button_filter in button_filter]
                    filters_data.append({sub_button_filters[0]: sub_button_filters[1:]})

        return filters_data

    def save_to_excel(self, data: list) -> None:
        """
        Сохраняет данные в файле Excel.
        Параметр data - список данных для сохранения.
        """
        self.ws.append(data)

        self.wb.save('data.xlsx')

    def data_retrieval(self, driver: webdriver.Chrome, link: WebElement) -> list:
        """
        Извлекает данные с отдельной страницы и возвращает список данных.
        Параметры:
        - driver: экземпляр WebDriver
        - link: элемент ссылки на страницу
        """
        href = link.find_element(By.TAG_NAME, 'a').get_attribute('href')
        driver.execute_script("window.open('" + href + "');")
        driver.switch_to.window(driver.window_handles[-1])

        if driver.find_element(By.TAG_NAME, 'h1').text == '429 Too Many Requests':
            time.sleep(1)
            driver.refresh()
        
        filters = self.get_filters(driver)

        try:
            count_offers = driver.find_element(By.CSS_SELECTOR, '.counter-item.active').find_element(By.CLASS_NAME, 'counter-value').text
        except NoSuchElementException:
            count_offers = None
            


        sell_button = driver.find_element(By.XPATH, "//a[contains(@class, 'btn-wide') and contains(text(), 'Продать')]")
        sell_button.click()

        self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        
        time.sleep(1)

        try:
            button_for_offer = driver.find_element(By.CLASS_NAME, 'js-lot-offer-edit')
        except NoSuchElementException:
            button_for_offer = None
        try:
            button_for_offer2 = driver.find_element(By.CSS_SELECTOR, '.tc.tc-selling')
        except NoSuchElementException:
            button_for_offer2 = None
            

        if button_for_offer:
            table_price = self.get_offer(driver, True)
        elif button_for_offer2:
            table_price = self.get_offer(driver, False)
        else:
            table_price = ['None']
        
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

        return [str(filters)] + [count_offers] + [*table_price]

    def get_offer(self, driver, section: bool) -> list:
        """
        Получает данные по ценам предложений.
        Возвращает список данных о ценах.
        """
        table_price = []

        # Если первая сенция, то используем node_id для запроса, иначе используется game
        if section:
            param = 'nodeId'
            id = driver.find_element(By.CLASS_NAME, 'js-lot-offer-edit').get_attribute('data-node')
            url = 'https://funpay.com/lots/calc'
        else:
            param = 'game'
            id = driver.find_element(By.NAME, 'game').get_attribute('value')
            url = 'https://funpay.com/chips/calc'

        for input_price in self.inputs_price:
            response = self.s.post(url=url, data={param: id, 'price': input_price}, headers=headers)
        
            if response.status_code != 200:
                time.sleep(1)
                response = self.s.post(url=url, data={param: id, 'price': input_price}, headers=headers)
            
            prices = jmespath.search("methods[].price", response.json())
            units = jmespath.search("methods[].unit", response.json())

            results = [f"{price} {unit}" for price, unit in zip(prices, units)]
            table_price.extend([input_price] + results)

        return table_price
    

if __name__ == '__main__':
    parser = FunPayParser()
    parser.parse('https://funpay.com/')
