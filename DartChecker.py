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
    ["시장구분", "종목코드", "종목명", "업종"])

driver = webdriver.Chrome(os.getcwd() + "\\chromedriver.exe")
wait = WebDriverWait(driver, 15)

df = pd.read_excel(os.getcwd() + '\\비상장 33135개 기업코드.xlsx', converters={'종목코드': str})
market = df['시장구분'].astype(str).values.tolist()
code = df['종목코드'].astype(str).values.tolist()
name = df['종목명'].astype(str).values.tolist()
line = df['업종'].astype(str).values.tolist()
data = zip(market, code, name, line)

start = time.time()
count = 0
count1 = 0
tmp_list = []
for a, b, c, d in data:
    count1 += 1
    driver.get('http://test.38.co.kr/forum2/dart.php?code=' + str(b))
    try:
        driver.find_element_by_xpath('//*[@id="report"]/table[2]/tbody/tr[2]/td[2]').text
        tmp_list = []
        count += 1
        print(count1, ' / ', count)
        print('종목코드 : ' + str(b))
        print('Check')
        tmp_list.append(a)
        tmp_list.append(str(b))
        tmp_list.append(c)
        tmp_list.append(d)
        print(tmp_list)
        excel_sheet.append(tmp_list)
        excel.save(filename='재무재표 보유 기업.xlsx')
    except:
        pass

print("time : ", time.time() - start)
# http://forum.38.co.kr/html/forum/board/?o=cinfo&code=366030
