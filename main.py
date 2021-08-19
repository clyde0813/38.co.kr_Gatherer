import os
import pandas as pd
import numpy as np
from openpyxl import Workbook

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
    ["시장구분", "종목코드", "종목명", "업종", "자본금", "수정(등록)일"])

driver = webdriver.Chrome(os.getcwd() + "\\chromedriver.exe")
wait = WebDriverWait(driver, 15)


def gatherer(n, m):
    for j in range(n, m):
        driver.get('http://www.38.co.kr/html/forum/com_list/?menu=nostock&page=' + str(j))
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[3]/tbody/tr/td/table[2]/tbody/tr/td/a[1]')))
        print('Page : ', str(j))
        table = driver.find_element_by_xpath('/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody')
        tr = table.find_elements_by_tag_name("tr")
        print(len(tr))
        for i in range(1, len(tr)):
            tmp_list = []
            if i % 2 == 0:
                pass
            else:
                market = driver.find_element_by_xpath(
                    '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[' + str(
                        i) + ']/td[1]/a').text
                code = driver.find_element_by_xpath(
                    '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[' + str(
                        i) + ']/td[2]').text
                name = driver.find_element_by_xpath(
                    '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[' + str(
                        i) + ']/td[3]/a').text
                line = driver.find_element_by_xpath(
                    '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[' + str(
                        i) + ']/td[4]').text
                money = driver.find_element_by_xpath(
                    '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[' + str(
                        i) + ']/td[5]').text
                date = driver.find_element_by_xpath(
                    '/html/body/table[3]/tbody/tr/td/table[1]/tbody/tr/td[1]/table[3]/tbody/tr[' + str(
                        i) + ']/td[6]').text
                tmp_list.append(market)
                tmp_list.append(code)
                tmp_list.append(name)
                tmp_list.append(line)
                tmp_list.append(money)
                tmp_list.append(date)
                print(market, ' ', code, ' ', name, ' ', line, ' ', money, ' ', date)
                excel_sheet.append(tmp_list)
        excel.save(filename=str(n) + '~' + str(m) + 'Data.xlsx')
        print('\n')


if __name__ == '__main__':
    first = Process(target=gatherer, args=(1, 1500))

    first.start()
    first.join()
