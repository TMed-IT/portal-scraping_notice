from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from time import sleep
import os
import datetime
import requests
import json

line_notify_api = 'https://notify-api.line.me/api/notify'
slack_webhook_url = 'https://hooks.slack.com/services/~~~'
LOGINID = ""
PASSWORD = ""

debug = False


def line_token():
    if debug:
        return ""
    else:
        return ""


def send_image_line_notify(path):
    headers = {'Authorization': f'Bearer {line_token()}'}
    data = {'message': '更新がありました'}
    files = {'imageFile': open(path, "rb")}
    requests.post(line_notify_api, headers=headers, data=data, files=files)


def log_success():
    data = {'text': 'ポータル確認済み',
            'username': 'ポータル更新通知'
            }
    requests.post(slack_webhook_url, data=json.dumps(data))


def log_error():
    data = {'text': 'スクリプトの実行中にエラーが発生しました',
            'username': 'ポータル更新通知'
            }
    requests.post(slack_webhook_url, data=json.dumps(data))


try:
    browser = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()))
    browser.get("http://ep.med.toho-u.ac.jp/")

    elem_username = browser.find_element(By.NAME, "MAILADDRESS")
    elem_password = browser.find_element(By.NAME, "LOGINPASS")
    elem_username.send_keys(LOGINID)
    elem_password.send_keys(PASSWORD)
    elem_submit = browser.find_element(
        By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[1]/form/p[1]/input")
    sleep(1)
    elem_submit.click()
    url_main = "https://ep.med.toho-u.ac.jp/default.asp?tb=1&ifkm=M1"
    browser.get(url_main)
    sleep(1)
    elem_table_contents = browser.find_elements(
        By.XPATH, '//*[@id="T1"]/tbody/tr')
    nothing_new = True
    PATH = os.path.join(os.path.dirname(__file__), ("ScreenShot.png"))

    for index in range(3, len(elem_table_contents)):
        elem_table_contents = browser.find_elements(
            By.XPATH, '//*[@id="T1"]/tbody/tr')
        elem_table_row = elem_table_contents[index].find_elements(
            By.TAG_NAME, "td")
        # 無効な行を回避
        if (len(elem_table_row) != 0):
            date_str = elem_table_row[3].text
            date = datetime.datetime.strptime(
                date_str, "%m/%d %H:%M").replace(year=datetime.datetime.now().year)
            # 最終更新が1時間より前の場合break
            if (date < (datetime.datetime.now() - datetime.timedelta(hours=1))):
                break
            elem_table_row[4].find_element(By.TAG_NAME, "a").click()
            elem_info = browser.find_element(
                By.XPATH, '//*[@id="contents"]/div/table[2]/tbody')
            print(elem_info.screenshot(PATH))
            send_image_line_notify(PATH)
            nothing_new = False
            browser.get(url_main)
    if not nothing_new:
        os.remove(PATH)
    log_success()
except:
    log_error()
