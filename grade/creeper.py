from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup
import re
import time
import sys
from openpyxl import Workbook
from selenium.webdriver.chrome.options import Options
LOGIN_URL = 'https://web085004.adm.ncyu.edu.tw/NewSite/login.aspx?Language=zh-TW'

select_items = {'1': '學期成績查詢'}


def creep(select):
    account, password = getAccountAndPassword()
    driver = webdriver.Chrome(options=setChromeOption())
    login(driver, account, password)
    changeModeToWindowMode(driver)
    jumpToGradeHtml(driver, select)
    output = parseGradeHtmlToData(driver)
    storeDataAndSave(output)
    driver.close()


def getAccountAndPassword():
    if len(sys.argv) != 3:
        print("invalid format,please type like python creeper.py 'your_username' 'your_password'")
        exit(-1)
    account = sys.argv[1]
    password = sys.argv[2]
    return account, password


def setChromeOption():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    return chrome_options


def login(driver, account, password):
    driver.get(LOGIN_URL)  # 進入登入網站
    act = driver.find_element_by_id("TbxAccountId")
    act.send_keys(account)
    pwd = driver.find_element_by_id("TbxPassword")
    pwd.send_keys(password)  # 輸入帳號密碼
    submit = driver.find_element_by_name("BtnPreLogin")
    submit.click()  # 登入
    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: driver.current_url != LOGIN_URL)  # 等待跳轉頁面


def changeModeToWindowMode(driver):
    mode = driver.find_element_by_id("BtnMode")
    mode.click()
    button = driver.find_element_by_xpath("/html/body/div[4]/div[3]/button[2]")
    button.click()
    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: driver.current_url != 'https://web085004.adm.ncyu.edu.tw/NewSite/Index1.aspx')  # 等待跳轉頁面


def jumpToGradeHtml(driver, select):
    application = driver.find_element_by_link_text(select_items[select])
    application.click()  # 按學期成績查詢
    driver.switch_to_window(driver.window_handles[-1])
    button = driver.find_element_by_id("btnOK")
    button.click()
    time.sleep(0.2)


def parseGradeHtmlToData(driver):
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    table = soup.find_all(id=re.compile('GridView1_ct*'))
    for i in range(len(table)):
        table[i] = table[i].text
    output = []
    for i in range(0, len(table), 5):
        if table[i+3] != "停修":
            output.append([table[i], table[i + 1], table[i + 2], table[i + 3], table[i + 4]])
    return output


def storeDataAndSave(output):
    wb = Workbook()
    sheet = wb['Sheet']
    for i in range(1, len(output) + 1):
        for j in range(1, len(output[i - 1]) + 1):
            sheet.cell(row=i, column=j, value=output[i - 1][j - 1])

    wb.save('grade.xlsx')


select = '1'
creep(select)
