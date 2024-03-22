import pandas as pd
import os
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from bs4 import BeautifulSoup

USERNAME = ''
PASSWORD = ''

LOGIN_URL = 'https://agent-bank.com/service/login?next=%2Fhome'

GSHEET_ID = '1a...'
SERVICEACCOUNT_JSON = './file.json'

# Windowsセレニウム起動

service = Service(executable_path=r'D:\chromedriver\chromedriver.exe')
chrome_options = Options()
chrome_options.add_experimental_option('detach', True)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Windowsセレニウム起動

driver.get(LOGIN_URL)  
sleep(1)

def get_content_by_header(header_text):
    try:
        dt_element = driver.find_element (By.XPATH, f'//dt[text()="{header_text}"]')
        dd_element = dt_element.find_element (By.XPATH, './following-sibling::dd')
        content_text = dd_element.get_attribute('innerHTML')
        soup = BeautifulSoup(content_text, 'html.parser')
        content_text = soup.get_text(separator='\n').strip()
        return content_text
    except Exception as e:
        return ''
    
def get_kinmuchi1():
    try:
        dt_element = driver.find_element (By.XPATH, f'//div[text()="勤務地1"]')
        dd_element = dt_element.find_element (By.XPATH, './following-sibling::div')
        content_text = dt_element.get_attribute('innerHTML') + dd_element.get_attribute('innerHTML')
        soup = BeautifulSoup(content_text, 'html.parser')
        content_text = soup.get_text(separator='\n').strip()
        return content_text
    except Exception as e:
        return ''

try:
    # ユーザーネーム及びパースワードでログイン
    el_username = driver.find_element(By.CSS_SELECTOR, "input[name='email']")
    el_username.send_keys(USERNAME)
    sleep(1)

    el_password = driver.find_element(By.CSS_SELECTOR, "input[name='password']")
    el_password.send_keys(PASSWORD)
    sleep(1)

    el_username.submit()
    sleep(1)

    print('ユーザー名、パスワードでログイン済')

    # グーグルスプレッドシートへアクセス
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICEACCOUNT_JSON, scope)

    gc = gspread.authorize(credentials)

    worksheet = gc.open_by_key(GSHEET_ID).worksheet('TEST')

    # グーグルスプレッドシートへアクセス

    rows = worksheet.get_all_values()

    csv_filename = 'data.csv'
    if os.path.exists(csv_filename):
        df = pd.read_csv(csv_filename)
    else:
        df = pd.DataFrame()

    for idx, row in enumerate(rows, start=1):
        if row[0] and not row[1]:
            current_row = []
            target_url = row[0]
            driver.get(target_url)
            print (f'対応中: {idx} : {target_url}')
            sleep(4)
            current_row.append(get_content_by_header('必須要件'))
            current_row.append(get_content_by_header('仕事内容'))
            current_row.append(get_content_by_header('仕事の醍醐味'))
            current_row.append(get_content_by_header('歓迎要件（活躍できる経験）'))
            current_row.append(get_content_by_header('募集職種'))
            current_row.append(get_content_by_header('想定年収'))
            current_row.append(get_content_by_header('予定募集人数'))
            current_row.append(get_content_by_header('管理監督者求人'))
            current_row.append(get_content_by_header('労働時間区分'))
            current_row.append(get_content_by_header('雇用形態'))
            current_row.append(get_content_by_header('勤務時間'))
            sleep(1)
            current_row.append(get_content_by_header('残業時間'))
            current_row.append(get_content_by_header('時短勤務'))
            current_row.append(get_content_by_header('選考詳細'))
            current_row.append(get_kinmuchi1())
            current_row.append(get_content_by_header('試用期間'))
            current_row.append(get_content_by_header('試用期間詳細'))
            current_row.append(get_content_by_header('給与・待遇'))
            current_row.append(get_content_by_header('年間休日'))
            current_row.append(get_content_by_header('休日・休暇'))
            current_row.append(get_content_by_header('福利厚生'))
            current_row.append(get_content_by_header('受動喫煙対策'))
            current_row.append(get_content_by_header('受動喫煙対策（詳細）'))
            current_row_df = pd.DataFrame([current_row])
            df = pd.concat([df, current_row_df], ignore_index=True)
            df.to_csv(csv_filename, index=False, header=False)
            print ("CSV updated.")
            sleep (2)

except Exception as e:
    print(f'例外エラー: {type(e).__name__} \n {str(e)}')

driver.quit()