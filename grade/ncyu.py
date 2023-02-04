import ddddocr
from selenium.webdriver.support.ui import WebDriverWait
import os
import base64

LOGIN_URL = 'https://web085004.adm.ncyu.edu.tw/NewSite/login.aspx?Language=zh-TW'
ocr = ddddocr.DdddOcr()
captcha_img_path = './Captcha.jpg'


def login(driver, account, password):
    driver.get(LOGIN_URL)  # 進入登入網站

    act = driver.find_element_by_id("TbxAccountId")
    act.send_keys(account)
    pwd = driver.find_element_by_id("TbxPassword")
    pwd.send_keys(password)  # 輸入帳號密碼

    captcha_result = getCaptchaResult(driver)
    cap = driver.find_element_by_id("TbxCaptcha")
    cap.send_keys(captcha_result)  # 輸入驗證碼

    submit = driver.find_element_by_name("BtnPreLogin")
    submit.click()  # 登入

    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: driver.current_url != LOGIN_URL)  # 等待跳轉頁面


def getCaptchaResult(driver):
    downloadCaptchaImg(driver)
    with open(captcha_img_path, 'rb') as f:
        img_bytes = f.read()

    res = ocr.classification(img_bytes)

    deleteCaptchaImg()
    return res


def downloadCaptchaImg(driver):
    x_path = '//*[@id="Image1"]'

    img_base64 = driver.execute_script("""
        var ele = arguments[0];
        var cnv = document.createElement('canvas');
        cnv.width = ele.width; cnv.height = ele.height;
        cnv.getContext('2d').drawImage(ele, 0, 0);
        return cnv.toDataURL('image/jpeg').substring(22);    
        """, driver.find_element_by_xpath(x_path))

    with open(captcha_img_path, 'wb') as image:
        image.write(base64.b64decode(img_base64))


def deleteCaptchaImg():
    os.remove(captcha_img_path)