import asyncio
import aiohttp
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By

import re

from datetime import datetime

from functions import *


URL = "https://www.viking.com.tw/en/category/Precision-Chip-Resistor-AR-Series/Resistor-ARBTC.html"

columns = ["Part Number", "Description", "Brand", "Series", "Resistance", "Tolerance", "Power",
               "TCR (ppm /℃)", "Size", "Package"]

columns_codes = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]


async def fetch(session, url):
    async with session.get(url) as response:
        return await response.text()


async def parse(start, end, filename, k):
    low = start
    async with aiohttp.ClientSession() as session:
        res = create_file(columns, columns_codes, filename)
        if res:
            low = res + 1  # установка первой страницы для парсинга

        driver = create_driver(URL)
        parsed_data = []
        for i in range(low, end + 1):
            url = URL+f'?page={i}'
            print(i)
            driver.get(url)
            html = await fetch(session, url)

            table = driver.find_elements(By.TAG_NAME, "table")[-1].find_element(By.TAG_NAME,
                                                                                "tbody")  # получение тела таблицы

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

            if i % 25 == 0 or i == end + 1:
                load_to_excel('viking_result.xlsx', parsed_data)
                parsed_data = []
                print(f"{i} страница закрыта в ", datetime.now().strftime("%H:%M:%S"))

            print(f"Поток {k}, кол-во элементов {len(parsed_data)}")

            load_to_excel(filename, parsed_data)
            parsed_data = []


async def main():
    tasks = []
    filename_template = 'viking_result_{}.xlsx'
    for k in range(4):
        start = k * 1000 + 1
        end = min((k + 1) * 1000, 3956)
        filename = filename_template.format(k + 1)
        task = asyncio.create_task(parse(start, end, filename, k))
        tasks.append(task)
    await asyncio.gather(*tasks)

if __name__ == '__main__':
    errors = 0
    print("Парсинг начался в ", datetime.now().strftime("%H:%M:%S"))
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        asyncio.run(main())
    except:
        errors += 1
    print("Парсинг закончился в ", datetime.now().strftime("%H:%M:%S"), "ошибок: ", errors)
