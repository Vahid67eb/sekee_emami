import os
import time
import requests
import pandas as pd
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from io import StringIO
import matplotlib.pyplot as plt

driver=webdriver.Chrome()
driver.get("https://www.google.com/")

def fetch_page_data(driver, page):
    url_template = 'https://www.tgju.org/profile/sekee/history?page={}'
    url = url_template.format(page)
    driver.get(url)
    time.sleep(2)  # اضافه کردن تاخیر برای بارگذاری کامل صفحه

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    table = soup.find('table')

    if table:
        data = pd.read_html(StringIO(str(table)))[0]
        return data
    else:
        print(f"No table found on page {page}")
        return None

def fetch_data(driver):
    all_data = []
    for page in range(1, 145):  # برای ۱۴۵ صفحه که شامل ۴۳۱۳ رکورد است
        data = fetch_page_data(driver, page)
        if data is not None:
            all_data.append(data)
        else:
            break

    combined_data = pd.concat(all_data, ignore_index=True)
    return combined_data

def save_to_excel(data):
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet.range('A1').value = data

    # ایجاد نمودار تغییرات درصدی
    date_column = 'تاریخ / شمسی'  # ستون تاریخ شمسی
    change_column = 'درصد تغییر'  # ستون درصد تغییر

    plt.figure(figsize=(10, 6))
    plt.plot(data[date_column], data[change_column], marker='o')
    plt.xlabel('تاریخ شمسی')
    plt