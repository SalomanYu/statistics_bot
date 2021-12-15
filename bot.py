"""
По сути данный бот представляет собой модель автоматизации бизнес-процесса.
В его задачи входит собирать информацию о заказах покупателей за предыдущий день,
высчитывать прибыль каждого товара, полагаясь на актуальную маржу из таблицы расчетов.
После нахождения актуальной маржи, нужно обновить таблицу со статистикой заказов товаров, заполнив
ячейки: количество товара, суммарная прибыль за день.


Алгоритм Работы бота:
I. Selenium 
    1. Авторизация на сайте https://online.moysklad.ru/
    2. Переход на необходимую вкладку и настройка фильтров
    3. Парсинг информации о заказе: Название организации, у которой был сделан заказ и артикул товара
    4. Запись спарсенных данных в отдельный файл для дальнейшей работы
II. Google Sheet API
    1. В зависимости от организации товара открываем нужную страницу таблицы (https://docs.google.com/spreadsheets/d/1bGbNieNgqDNSORaphLhLOHUbIUE00yxA0q_b4HsNclM/edit#gid=1655640849)
    2. Поиск по таблице расчетов  актуальной маржи товара
    3. Рассчитываем прибыль конткретного товара за день по формуле: кол-во заказов * маржа товара = прибыль от товара за день
    4. Записываем в таблицу статистики прибыль от товара за день и количество заказов (https://docs.google.com/spreadsheets/d/1rEGdqDGFzdaSAlTzjiFt-GlW-scgLx2-UDgQdN0PL_s/edit#gid=1225586096)
    5. Завершаем работу


Дополнительная информация:
    1. В боте используется система хранения историй работы бота, т.е создается директория с датой запуска бота,
        куда будут сохранены файлы спарсенных данных и актуальной маржи товаров.
        В дальнейшем можно установить срок хранения истории (очищать ее через 3/6/9/12 месяцев)

    2. Данный бот в текущей версии настроен на работу лишь с двумя организациями ИП Ермалович и ИП Александров.
        Пока это имеет огромное значение, так как в некоторыъ блоках кода явно указываются эти организации
        То есть в лпане работы с организациями бот ПОКА не универсален. 
"""




from time import sleep
import datetime, os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import gspread
from oauth2client.service_account import ServiceAccountCredentials  


tomorrow = str((datetime.date.today() - datetime.timedelta(days=1)).day) # переменная для нахождения колонки вчерашнего дня в таблице статистики
tomorrow_date = str(datetime.date.today() - datetime.timedelta(days=1)) # переменная для создания директорий
parent_directory = os.getcwd()
history_directory = os.path.join(parent_directory,tomorrow_date)

os.makedirs(history_directory, exist_ok=True) # Создаем директории истории за текущий день, если она не была создана 

# Декораторы для красивого и читабельного вывода в консоль, подсвечивают соответствующую информацию о состоянии работы бота
success_message = '\033[2;30;42m [SUCCESS] \033[0;0m' 
warning_message = '\033[2;30;43m [WARNING] \033[0;0m'

result_parser = f"{history_directory}/parser_result.txt"

if result_parser:
   os.remove(result_parser) 


class SeleniumParser: 
    
    """
    Этот класс выполняет первый блок вышеописанного алгоритма
    """

    def __init__(self, mysklag_login, mysklag_password):
        self.password_user = mysklag_password
        self.login_user = mysklag_login


    def start(self):

        """
        Этот метод подключает Chrome Webdriver вместе с необходимыми настройками(опциями) и подключается к url
        """

        print(warning_message + '\tЗапустили бота...')
        url = 'https://online.moysklad.ru/'
        print(success_message + '\tОткрыли сайт ', url)

        option = Options()
        
        option.add_argument("--headless") # ФОНОВЫЙ РЕЖИМ   
        # Отключаем всплывающие сообщения и окна браузера
        option.add_argument("--disable-infobars") 
        option.add_argument("start-maximized")
        option.add_argument("--disable-extensions")
        option.add_experimental_option("prefs", { 
            "profile.default_content_setting_values.notifications": 2
        })


        self.browser = webdriver.Chrome(options=option)
        self.browser.maximize_window()
        self.browser.get(url)    
        sleep(2)

        self.authorize(self.password_user, self.login_user)

    def authorize(self, user_password, user_login):

        """
        Метод авторизации на сайте 
        """

        login = self.browser.find_element(By.ID, 'lable-login')
        login.send_keys(user_login)

        password = self.browser.find_element(By.ID, 'lable-password')
        password.send_keys(user_password)

        button = self.browser.find_element(By.CLASS_NAME, 'b-button')
        button.click()
        print(success_message + '\tВошли в учетную запись...')
        sleep(6)

        self.open_current_table()


    def open_current_table(self):

        """
        настройка фильтров перед парсингом информации о заказах
        """

        button_sales = self.browser.find_element(By.XPATH, "//div[@class='topMenu-new']//td[@class='topMenuItem-new'][2]")
        button_sales.click()
        sleep(1)

        button_orders = self.browser.find_element(By.XPATH, "//div[@class='subMenuContainer-new']//span")
        button_orders.click()
        print(success_message + "\tОткрыли таблицу с заказами покупателей за вчерашнее число...")
        sleep(3)

        button_period = self.browser.find_element(By.XPATH, "//div[@class='period-filter-widget2-preset-label']")
        button_period.click()
        sleep(1)

        button_refresh = self.browser.find_element(By.CLASS_NAME, "b-tool-button")
        button_refresh.click()
        sleep(1)
        self.parse_data()
        while True:
            try:
                next_page = self.browser.find_element_by_xpath("//td[@class='next-page']//img[@class='gwt-Image']")
                next_page.click()
                sleep(3)
                self.parse_data()

            except BaseException:
                break                

    def parse_data(self):

        """
        Нюанс данного метода заключается в том, что строки в таблице поделены
        на четные и нечетные и поэтому приходится парсить обе строки по разному
        После чего объединять всё воедино
        """

        even_organizations = self.browser.find_elements(By.XPATH, "//td[@class='cellTableCell cellTableEvenRowCell '][3]")
        odd_organizations = self.browser.find_elements(By.XPATH, "//td[@class='cellTableCell cellTableOddRowCell '][3]")


        all_organizations = even_organizations + odd_organizations
        text_organizations = [item.text for item in all_organizations]

        even_comments = self.browser.find_elements(By.XPATH, "//td[@class='cellTableCell cellTableEvenRowCell '][7]")
        odd_comments = self.browser.find_elements(By.XPATH, "//td[@class='cellTableCell cellTableOddRowCell '][7]")

        print(success_message + '\tСпарсили название организаций и комментарии к заказам...')   

        all_comments = even_comments + odd_comments
        text_comments  = [item.text for item in all_comments]

        id_orders = [item.split(',')[0] for item in text_comments]

        print(success_message + '\tСохраняем собранные данные...')

        self.save_data(text_organizations, id_orders)


    def save_data(self, organizations, ids):
        
        """
        Сохраняем спарсенные данные в файлик parser_result.txt
        """

        filename = f'{history_directory}/parser_result.txt'

        with open(filename, 'a') as file:
            for item in range(len(organizations)):
                file.write(f'{organizations[item]} - {ids[item]} \n')

        print(success_message + '\tЗаписали файл ' + filename)


    def get_frequency_dict(self):

        """
        Создаем частотный словарь
        """

        organizations_with_orders = []
        with open(f'{history_directory}/parser_result.txt', 'r') as file:
            for line in file:
                organization = line.split('-')[0]
                order = line.split('-')[1].replace("\n", '').strip()

                organizations_with_orders.append((organization, order))


        frequency_dictionary = {}
        for item in organizations_with_orders:
            count = 0
            for item2 in organizations_with_orders:
                if item[1] == item2[1]:
                    count += 1
            frequency_dictionary[item[1]] = count

        print(success_message + '\tОтсортировали собранные данные в частотный словарь')
        return frequency_dictionary


class Spreadsheet:
    """
    Этот класс выполняет второй блок вышеописанного алгоритма
    """

    def run(self, frequency_dictionary):

        """
        Основной метод данного класса, объединяющий все остальные методы
        """

        spread = self.auth_spread('1bGbNieNgqDNSORaphLhLOHUbIUE00yxA0q_b4HsNclM') # инициализация таблицы 

        print(success_message + '\tПодключились к таблице расчетов')
        first_org_margins = self.get_margin_by_organization(spread, 'Александров А.А', frequency_dictionary) # Сбор маржи определенной организации
        
        print(warning_message + '\tБот взял паузу на одну минуту, чтобы избежать лимита на количество запросов в минуту.')
        sleep(80) # Чтобы обойти лимит по количеству запросов Google API

        second_org_margins = self.get_margin_by_organization(spread, "ИП Ермалович А.С", frequency_dictionary) # Сбор маржи определенной организации
        self.save_result(first_org_margins, second_org_margins)


    def auth_spread(self, spread_id):
        
        """
        Данный метод отвечает за подключение к таблице 
        """


        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('rosya-project-91e7c34fdc81.json', scope)

        gc = gspread.authorize(credentials)
        spread = gc.open_by_key(spread_id)

        return spread

    def get_margin_by_organization(self, spread, organization, frequency_dictionary):

        """
        СОбираем маржу и высчитываем общую прибыль
        """

        worksheet = self.open_worksheet(spread, organization) # открываем нужную страницу, полагаясь на название организации
        col_margin = worksheet.find('Маржа').col
        organization_margin_orders = []
        
        for order in frequency_dictionary:
            row_order = worksheet.find(str(order))
            if row_order != None:
                margin_order = worksheet.cell(row_order.row, col_margin).value
                amount_margin_order = float(margin_order.split('₽')[0].replace(',', '.')) * frequency_dictionary[order]
                organization_margin_orders.append((order, round(amount_margin_order, 2), frequency_dictionary[order]))

        print(success_message + '\tСобрали маржу организации: ', organization)
        return organization_margin_orders

    def open_worksheet(self, spread, worksheet_name):

        """
        САМЫЙ НЕУДАЧНЫЙ МЕТОД, КОТОРЫЙ В ОБЯЗАТЕЛЬНОМ ПОРЯДКЕ НУЖНО ИЗМЕНИТЬ НА ЧТО-ТО УНИВЕРСАЛЬНОЕ
        """

        if worksheet_name == 'Александров А.А':
            return spread.get_worksheet(0)

        elif worksheet_name == "ИП Ермалович А.С":
            return spread.get_worksheet(1)
    


    def save_result(self, first_org, second_org):

        """
        сохраняем результат для истории
        """

        filename = f'{history_directory}/margin_orders.txt'

        with open(filename, 'w') as file:
            for item in first_org:
                file.write(f'{item[0]} - {item[1]}  - {item[2]} \n')

            for item in second_org:
                file.write(f'{item[0]} - {item[1]}  - {item[2]} \n')

        print(success_message + '\tЗаписали файл ' + filename)
        self.update_statistics_table()


    def update_statistics_table(self):

        """
        Обновляем информацию о статистике
        """

        spread = self.auth_spread('1rEGdqDGFzdaSAlTzjiFt-GlW-scgLx2-UDgQdN0PL_s')
        worksheet = spread.worksheet('Статистика (Декабрь)')
        orders = []
        margins = []

        # ФАЙЛ ПОКА НЕ МОЖЕТ РАБОТАТЬ С ЛИСТАМИ ТАБЛИЦЫ 
        print(warning_message + '\tБот взял паузу на одну минуту, чтобы избежать лимита на количество запросов в минуту.')
        sleep(80) # Чтобы обойти лимит по количеству запросов Google API
        
        print(warning_message + '\tОбновляем статистику')

        with open(f'{history_directory}/margin_orders.txt', 'r') as file:
            for line in file:
                order = line.split('-')[0].strip()
                margin = line.split('-')[1].replace('\n', '').strip().replace('.', ',')
                count = line.split('-')[2].strip()
                

                order_row = worksheet.find(order).row
                order_count_row = order_row - 1
                tomorrow_col = worksheet.find(tomorrow).col

                worksheet.update_cell(order_row, tomorrow_col, margin)
                worksheet.update_cell(order_count_row, tomorrow_col, count)
        print(success_message + '\tБот успешно завершил свою работу')




bot_selenium = SeleniumParser(mysklag_login='vika@ermalovich1972', mysklag_password='Ugegeg')
bot_selenium.start()
bot_selenium.browser.quit()

frequen_dict = bot_selenium.get_frequency_dict()

spread = Spreadsheet()
spread.run(frequen_dict)
