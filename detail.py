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
    ["시장구분", "종목코드", "회사명", "업종", "홈페이지", "대표이사", "대표전화", "팩스", "본사주소", "설립일", "사업자 번호", "결산원", "직원수", "보통주", "액면가",
     "우선주",
     "자본금",
     "전체 주식수", "주권구분", "주거래은행", "명의개서 여부", "계좌이체 여부", "대행기관", "주식담당", "주식문의"])

driver = webdriver.Chrome(os.getcwd() + "\\chromedriver.exe")
wait = WebDriverWait(driver, 15)

df = pd.read_excel(os.getcwd() + '\\1~31Data.xlsx', converters={'종목코드': str})
df = df['종목코드'].astype(str).values.tolist()
print(df)

start = time.time()

tmp_list = []
for i in df:
    tmp_list = []
    print('종목코드 : ' + str(i))
    driver.get('http://forum.38.co.kr/html/forum/board/?o=cinfo&code=' + i)
    wait.until(EC.presence_of_element_located((By.XPATH,
                                               '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[1]/td[2]')))
    company_name = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[1]/td[2]').text
    line = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[1]/td[4]').text
    website = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[2]/td[2]').text
    owner = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[2]/td[4]').text
    phone = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[3]/td[2]').text
    fax = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[3]/td[4]').text
    location = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[4]/td[2]').text
    create_date = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[5]/td[2]').text
    own_num = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[5]/td[4]').text
    last_month = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[6]/td[2]').text
    employee = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[4]/td/table/tbody/tr[6]/td[4]').text

    ordinary = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[1]/td[2]').text
    price = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[1]/td[4]').text
    prefer = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[2]/td[2]').text
    own_money = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[2]/td[4]').text
    total_stock = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[3]/td[2]').text
    classification = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[3]/td[4]').text
    bank = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[4]/td[2]').text
    name_classify = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[4]/td[4]').text
    transfer = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[5]/td[2]').text
    agency = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[5]/td[4]').text
    manager = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[6]/td[2]').text
    ask = driver.find_element_by_xpath(
        '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[7]/td/table/tbody/tr[6]/td[4]').text
    tmp_list.append('비상장')
    tmp_list.append(str(i))
    tmp_list.append(company_name)
    tmp_list.append(line)
    tmp_list.append(website)
    tmp_list.append(owner)
    tmp_list.append(phone)
    tmp_list.append(fax)
    tmp_list.append(location)
    tmp_list.append(create_date)
    tmp_list.append(own_num)
    tmp_list.append(last_month)
    tmp_list.append(employee)
    tmp_list.append(ordinary)
    tmp_list.append(price)
    tmp_list.append(prefer)
    tmp_list.append(own_money)
    tmp_list.append(total_stock)
    tmp_list.append(classification)
    tmp_list.append(bank)
    tmp_list.append(name_classify)
    tmp_list.append(transfer)
    tmp_list.append(agency)
    tmp_list.append(manager)
    tmp_list.append(ask)
    print(tmp_list)
    excel_sheet.append(tmp_list)
    excel.save(filename='기업개요.xlsx')

print("time : ", time.time() - start)
# http://forum.38.co.kr/html/forum/board/?o=cinfo&code=366030
