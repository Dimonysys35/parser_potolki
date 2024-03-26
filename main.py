from bs4 import BeautifulSoup
import requests
import pandas as pd
import telebot
import os
from telebot import types
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# 5378837779:AAElD5McdczG-ezyU4CqcmWomX_O1aR0ads

bot = telebot.TeleBot('5378837779:AAElD5McdczG-ezyU4CqcmWomX_O1aR0ads')


@bot.message_handler(commands=['start'])
def hi(message):
    bot.send_message(message.chat.id, call(message))


@bot.message_handler(commands=['callback'])
def call(message):
    k_b = types.InlineKeyboardMarkup(row_width=2)
    my_like_treks = types.InlineKeyboardButton(text='POTOLKI-SIMFEROPOL', callback_data='mag1')
    my_like_treks2 = types.InlineKeyboardButton(text='STYLISHROOM', callback_data='mag2')
    my_like_treks3 = types.InlineKeyboardButton(text='SVD_POTOLKI', callback_data='mag3')
    my_like_treks4 = types.InlineKeyboardButton(text='GORIZONT_KRIM', callback_data='mag4')
    k_b.add(my_like_treks, my_like_treks2, my_like_treks3, my_like_treks4)
    bot.send_message(message.chat.id, "Данный бот парсит выбранный вами сайт, по продаже натяжных потолков, и присылает EXCEL файл с данными.\nВыберите магазин:", reply_markup=k_b)

@bot.callback_query_handler(func=lambda c: c.data)
def ancwer(callback):
    # POTOLKI-SIMFEROPOL
    if callback.data == 'mag1':
        url = 'https://potolok-simferopol.ru/ceni'
        headers = {
            'accept': '*/*',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.63 Safari/537.36'
        }

        req = requests.get(url, headers=headers)
        src = req.text

        with open('index.html', 'w', encoding="utf-8") as file:
            file.write(src)

        with open('index.html', encoding="utf-8") as file:
            src = file.read()

        soup = BeautifulSoup(src, 'lxml')
        df = pd.DataFrame(columns=["Название", "Цена"])
        all_cards = soup.find_all(class_='s-elements-grid valign-top use-flex')
        for card in all_cards:
            first_card_lst = card.find_all(class_='cont cell')
            for el in first_card_lst:
                try:
                    text = el.find('div').attrs['data-item']
                    data = eval(text)
                    row = [data[0]["value"], data[1]["value"]]
                    #print(row)
                    length = len(df)
                    df.loc[length] = row
                    #print(f'\n{data[0]["name"]}: {data[0]["value"]}\n{data[1]["name"]}: {data[1]["value"]}')
                except KeyError:
                    break
        df.to_excel('potolok-simferopol.xlsx', sheet_name="Расценки potolok-simferopol.ru", index=False)
        bot.send_document(callback.message.chat.id, open(r'potolok-simferopol.xlsx', 'rb'))
        print("potolok-simferopol.xlsx")
        #os.remove('potolok-simferopol.xlsx')
    # STYLISHROOM
    if callback.data == 'mag2':
        # получение страницы через get запрос
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.63 Safari/537.36"
        }
        response = requests.get('https://stylishroom.com.ru/stretch-ceiling/price/', headers=headers)
        soup = BeautifulSoup(response.text, "lxml")
        # название таблицы
        tlines = soup.find_all("div", class_="tline")
        # таблицы и словарь для хранения данных
        tables = soup.find_all("table", cellpadding="8")
        data = {}
        cnt = 0
        # самая важная часть - сбор данных их таблиц
        for table in tables:
            # headers - это названия столбцов
            preheaders = []
            for i in table.find_all("tr", bgcolor="#EDEBD5"):
                title = i.text
                if title != "\n":
                    preheaders.append(title.split("\n"))
            for i in preheaders:
                for j in i:
                    if j == "":
                        i.pop(i.index(j))
            if len(preheaders) > 1:
                preheaders[0].pop(len(preheaders[0]) - 1)
                headers = preheaders[0] + preheaders[1]
            else:
                headers = preheaders[0]
            for i in headers:
                if i == "":
                    headers.pop(headers.index(i))
            # tr - это строки, mydata - это датафрейм для каждой таблицы с названием колонок как у хедеров
            trs = table.find_all("tr", bgcolor="#ffffff")
            mydata = pd.DataFrame(columns=headers)
            # каждая строка разбирается отдельно
            for tr in trs:
                # получение текста из ячеек и сбор их в список(очень слабая часть, тк, если текст будет обернут в
                # другие теги(ul, li, a, другие div) или в таблице будет картика, то код сломается)
                row_data = tr.find_all("td")
                row = [i.text for i in row_data]
                length = len(mydata)
                # если в списке-строке недостаточно элементов, то нужно добавить "None" в начало списка !!!! костыль !!!!
                while len(row) < len(headers):
                    row.insert(0, "None")
                # добавление строки в датафрейм
                mydata.loc[length] = row
            # сохранение таблицы в словаре
            data[tlines[cnt].text] = mydata
            cnt += 1
        # создание эксель файла через pandas
        cnt = 0
        writer = pd.ExcelWriter("stylishroom.xlsx", engine="xlsxwriter")
        for i in data:
            name = ""
            cnt = 0
            for c in i:
                if cnt == 28:
                    break
                if c in ["[", "]", ":", "*", "?", "/", "\\"]:
                    name += ""
                else:
                    name += c
                cnt += 1
            data[i].to_excel(writer, sheet_name=name, index=False)
        writer.save()
        bot.send_document(callback.message.chat.id, open(r'stylishroom.xlsx', 'rb'))
        print("stylishroom.xlsx")
    #SVD_POTOLKI
    '''для работы этого парсера нужен вебдрайвер для хрома'''
    if callback.data == "mag3":
        # добавление опций для драйвера
        options = webdriver.ChromeOptions()
        options.add_argument("user-agent=Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:84.0) Gecko/20100101 Firefox/84.0")
        options.add_argument("--disable-blink-features=AutomationControlled")
        # headless mode
        # options.add_argument("--headless")
        options.headless = True
        # подключаемся к сайту
        url = "https://svd-potolki.ru/price"
        driver = webdriver.Chrome(executable_path="chromedriver.exe", options=options)
        driver.get(url)
        # установка площади и количества люстр
        calc_area = driver.find_element(By.ID, "calc-area")
        calc_area.send_keys("1")
        calc_lamp = driver.find_element(By.ID, "calc-lamp")
        calc_lamp.send_keys("1")
        # открытие блока с select-ами
        choose = driver.find_element(By.CLASS_NAME, "open-facture")
        choose.click()
        # колонки, начало строки и создание датафрейма
        columns = ["Площадь(м^2)", "Периметр", "Углы", "Светильники", "Фактура", "Цвет", "Ширина(м)", "Производитель",
                   "Цена", "Гарантия"]
        start = [1, 0, 4, 1]
        data = pd.DataFrame(columns=columns)
        # поиск первого select
        facture = Select(driver.find_element(By.XPATH, "//*[@id=\"calc-facture\"]"))
        for i in facture.options:
            # нажатие на опцию
            i.click()
            # очистка и втавка площади потолка(костыль), для того чтобы данные в нужном блоке обновились
            calc_area.clear()
            calc_area.send_keys("1")
            # поиск следующего выпадающего списка
            makers = Select(driver.find_element(By.ID, "calc-producer"))
            # тоже самое для следующего списка
            for j in makers.options:
                j.click()
                calc_area.clear()
                calc_area.send_keys("1")
                colors = Select(driver.find_element(By.ID, "calc-color"))
                for g in colors.options:
                    g.click()
                    calc_area.clear()
                    calc_area.send_keys("1")
                    width = Select(driver.find_element(By.ID, "calc-width"))
                    for k in width.options:
                        k.click()
                        calc_area.clear()
                        calc_area.send_keys("1")
                        # создание супа
                        soup = BeautifulSoup(driver.page_source, 'lxml')
                        # поиск блока с информацией
                        info = soup.find("div", "info")
                        row = []
                        row += start
                        length = len(data)
                        # выборка информации из каждого параграфа в блоке
                        for p in info:
                            if p.text != "\n":
                                text = p.text.split(':')
                                # print(text)
                                text = text[1].replace("\n", "")
                                # print(p, "---------", text)
                                row.append(text)
                        # print(row, "++-+-+-++")
                        # заполнение недостающих ячеек в строке(костыль)
                        while len(row) < len(columns):
                            # print(row)
                            row.insert(0, "None")
                        # добавление строки в дф
                        data.loc[length] = row
        # создание эксель файла
        writer = pd.ExcelWriter("svd_potolki.xlsx", engine="xlsxwriter")
        data.to_excel(writer, sheet_name="svd-potolki(калькулятор)", index=False)
        writer.save()
        driver.quit()
        bot.send_document(callback.message.chat.id, open(r'svd_potolki.xlsx', 'rb'))
        print("svd_potolki.xlsx")
    #GORIZONT_KRIM
    if callback.data == "mag4":
        r = requests.get('https://gorizont-krim.ru/pricelist/')
        soup = BeautifulSoup(r.content, 'lxml')
        sheet_names = ["ПВХ", "Тканевые", "Сатиновые", "Белые", "Цветные", "Двухуровневые", "Фотопечать", ""]
        tables = soup.find_all('table')
        tables.pop(0)
        tables.pop(len(tables) - 1)
        data = {}
        cnt = 0
        for table in tables:
            headers = []
            for i in table.find_all("thead"):
                for th in table.find_all("th"):
                    title = th.text
                    if title != "\n":
                        headers.append(title)
            mydata = pd.DataFrame(columns=headers)
            for tr in table.find("tbody").find_all("tr"):
                at = tr.attrs
                if len(at) == 0:
                    clas = ""
                else:
                    clas = at["class"][0]
                if clas != "subname" or clas == "" or clas == "grey_tr":
                    row = []
                    for th in tr.find_all("td"):
                        text = th.text
                        row.append(text)
                    length = len(mydata)
                    while len(row) < len(headers):
                        row.insert(0, "None")
                    mydata.loc[length] = row
            data[sheet_names[cnt]] = mydata
            cnt += 1
        writer = pd.ExcelWriter("gorizont-Krim.xlsx", engine="xlsxwriter")
        for i in data:
            name = ""
            cnt = 0
            for c in i:
                if cnt == 28:
                    break
                if c in ["[", "]", ":", "*", "?", "/", "\\"]:
                    name += ""
                else:
                    name += c
                cnt += 1
            data[i].to_excel(writer, sheet_name=name, index=False)
        writer.save()
        bot.send_document(callback.message.chat.id, open(r'gorizont-Krim.xlsx', 'rb'))
        print("gorizont-Krim.xlsx")
bot.infinity_polling()