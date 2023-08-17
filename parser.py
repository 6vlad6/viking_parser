from selenium import webdriver
from selenium.webdriver.common.by import By

import re

from datetime import datetime

from functions import *


URL = "https://www.viking.com.tw/en/category/Precision-Chip-Resistor-AR-Series/Resistor-ARBTC.html"
MIN_PAGE = 1
MAX_PAGE = 1


columns = ["Part Number", "Description", "Brand", "Series", "Resistance", "Tolerance", "Power",
               "TCR (ppm /℃)", "Size", "Package"]

columns_codes = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]

res = create_file(columns, columns_codes, "viking_result.xlsx")  # создание результирующего файла

if res:
    MIN_PAGE = res + 1  # установка первой страницы для парсинга


# поиск последней страницы
driver = create_driver(URL)

pagination_data = driver.find_element(By.CLASS_NAME, "control-post")
MAX_PAGE = int(pagination_data.find_elements(By.TAG_NAME, "a")[-2].text)

errors = 0
try:
    print("Парсинг начался в ", datetime.now().strftime("%H:%M:%S"))
    parsed_data = []
    for i in range(MIN_PAGE, MAX_PAGE+1):

        driver.get(URL+f"?page={i}")

        table = driver.find_elements(By.TAG_NAME, "table")[-1].find_element(By.TAG_NAME, "tbody")  # получение тела таблицы

        table_rows = table.find_elements(By.TAG_NAME, "tr")  # все строки таблицы
        for tr in table_rows:
            tds = tr.find_elements(By.TAG_NAME, "td")  # все столбцы

            full_part_number = tds[0].text
            part_number = ""

            match = re.search(r'AR Series\s+(.*?)\)', full_part_number)
            if match:
                part_number = match.group(1)

            size = tds[1].text
            tolerance = tds[2].text
            tcr = tds[3].text
            power = tds[4].text
            resistance = tds[5].text
            package = tds[6].text

            parsed_data.append([part_number, "Thin Film Precision Chip Resistor", "Viking", "AR",
                                                 resistance, tolerance, power, tcr, size, package])

        if i % 25 == 0 or i == MAX_PAGE+1:
            load_to_excel('viking_result.xlsx', parsed_data)
            parsed_data = []
            print(f"{i} страница закрыта в ", datetime.now().strftime("%H:%M:%S"))
except:
    errors += 1

print("Работа завершена в ", datetime.now().strftime("%H:%M:%S"), "ошибок: ", errors)