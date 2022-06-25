from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from win32com.client import Dispatch
from selenium.webdriver.chrome.options import Options
import re
import pandas as pd
import time
import os
import argparse
LOGIN_URL = 'https://web085004.adm.ncyu.edu.tw/NewSite/login.aspx?Language=zh-TW'

select_items = {'1': '學期成績查詢'}

is_update = False
args = {}
file_path = ""


def parseArgs():
    global args, file_path
    parser = argparse.ArgumentParser()
    parser.add_argument("account",
                        type=str,
                        help="請輸入帳號")
    parser.add_argument("password",
                        type=str,
                        help="請輸入密碼")
    parser.add_argument("-o",
                        nargs='?',
                        type=str,
                        default='.',
                        help="輸出的資料夾")
    parser.add_argument("-n",
                        nargs='?',
                        type=str,
                        default='grade',
                        help="輸出的檔案名稱")
    args = vars(parser.parse_args())
    if not os.path.isdir(args['o']):
        raise ValueError('輸出的路徑不是資料夾')

    if args['o'] == '.':
        file_path = os.path.join(os.path.dirname(__file__), args['n']+'.xlsx')
    else:
        file_path = os.path.join(args['o'], args['n']+'.xlsx')


def creep(select):
    account, password = getAccountAndPassword()
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=setChromeOption())
    login(driver, account, password)
    changeModeToWindowMode(driver)
    jumpToGradeHtml(driver, select)
    data = pd.DataFrame(parseHtmlToData(driver))
    if isUpdate(data):
        storeDataAndSave(data)
    driver.close()


def getAccountAndPassword():
    account = args['account']
    password = args['password']
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


def parseHtmlToData(driver):
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    output = parseGrade(soup)
    if soup.find(id='FVSelstchf_btnShowRank'):
        button = driver.find_element_by_id('FVSelstchf_btnShowRank')
        button.click()
        wait = WebDriverWait(driver, 10)
        wait.until(lambda driver: 'FVSelstchf_lblRank' in driver.page_source)  # 等待名次顯現

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        rank = parseRank(soup)
        if rank:
            output.append(rank)
    return output


def parseGrade(soup):
    table = soup.find_all(id=re.compile('GridView1_ct*'))
    for i in range(len(table)):
        table[i] = table[i].text
    output = []
    for i in range(0, len(table), 5):  # 成績
        if table[i + 3] != "停修":
            output.append([table[i], table[i + 1], table[i + 2], table[i + 3], table[i + 4]])
    return output


def parseRank(soup):
    rank_msg = soup.find(id='FVSelstchf_lblRank').text
    if rank_msg == '無名次':
        return []
    else:
        rank = ''
        for i in rank_msg:
            if i.isdigit():
                rank += i
        rank = int(rank)
        return ['名次', rank]


def isUpdate(data):
    return isDataUpdate(data) or isGradeUpdate()


def isDataUpdate(data):
    global is_update, file_path

    if os.path.isfile(file_path):
        old_data = pd.read_excel(file_path, dtype={0: str, 3: str}).iloc[:, :4]
        new_data = data.iloc[:, :4]
        if old_data.equals(new_data):
            return False
        else:
            is_update = True
            return True
    else:
        is_update = True
        return True


def isGradeUpdate():
    global file_path
    if os.path.isfile(file_path):
        old_data = pd.read_excel(file_path, dtype={0: str, 3: str}).iloc[:, :4]
        if old_data.iloc[-1][0] != '名次':
            return True
    return False


def storeDataAndSave(output):
    global file_path
    output.to_excel(file_path, index=False)


if __name__ == "__main__":
    try:
        parseArgs()
    except ValueError as e:
        print(e)
        exit(-1)

    select = '1'
    creep(select)
    if is_update:
        xl = Dispatch("Excel.Application")
        xl.Visible = True
        wb = xl.Workbooks.Open(file_path)
