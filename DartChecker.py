import os
import pandas as pd
import numpy as np
from openpyxl import Workbook
import time
from multiprocessing import Process

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

excel = Workbook()
excel_sheet = excel.create_sheet('종목코드')
excel_sheet = excel.active
excel_sheet.append(
    ["시장구분", "종목코드", "회사명", "업종"])

driver = webdriver.Chrome(os.getcwd() + "\\chromedriver.exe")
wait = WebDriverWait(driver, 15)

df = pd.read_excel(os.getcwd() + '\\비상장 33135개 기업코드.xlsx', converters={'종목코드': str})
df = df['종목코드'].astype(str).values.tolist()

start = time.time()
count = 0
count1 = 0
tmp_list = []
for i in df:
    count1 += 1
    tmp_list = []
    driver.get('http://test.38.co.kr/forum2/dart.php?code=' + i)
    try:
        driver.find_element_by_xpath('//*[@id="report"]/table[2]/tbody/tr[2]/td[2]').text
        count += 1
        print(count1, ' / ', count)
        print('종목코드 : ' + str(i))
        print('Check')
    except:
        pass

    # print(tmp_list)
    # excel_sheet.append(tmp_list)
    # excel.save(filename='기업개요.xlsx')

print("time : ", time.time() - start)
# http://forum.38.co.kr/html/forum/board/?o=cinfo&code=366030
