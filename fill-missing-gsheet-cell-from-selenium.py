from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from google_service_util import GoogleServiceUtil
import gspread
from oauth2client.service_account import ServiceAccountCredentials

USERNAME = ''
PASSWORD = ''

LOGIN_URL = 'https://agent-bank.com/service/login?next=%2Fhome'

GSHEET_ID = '1a...'
SERVICEACCOUNT_JSON = './file.json'

# Windowsセレニウム起動

service = Service(executable_path=r'D:\chromedriver\chromedriver.exe')
chrome_options = Options()
driver = webdriver.Chrome(service=service, options=chrome_options)

# Windowsセレニウム起動

driver.get(LOGIN_URL)  
sleep(2)

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

def get_kaishamei():
    try:
        el_kaishamei = driver.find_element (By.CSS_SELECTOR, ".company > .name")
        content_text = el_kaishamei.get_attribute('innerHTML');
        soup = BeautifulSoup(content_text, 'html.parser')
        content_text = soup.get_text(separator='\n').strip()
        return content_text;
    except Exception as e:
        return ''

def get_shokushumei():
    try:
        el_shokushumei = driver.find_element (By.CSS_SELECTOR, ".title > p")
        return el_shokushumei.get_attribute('textContent');
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

    rows = worksheet.get_values('A1:Z3636')
    headers = worksheet.row_values(1)

    for idx, row in enumerate(rows, start=1):
        if row[0]:  # 1番目のセールにURLが入っている場合
            if any(cell == '' for cell in row[3:24]):
                target_url = row[0]
                driver.get(target_url)
                print (f'対応中: {idx} : {target_url}')
                sleep(3)

                for col_index in range(3, 24):
                    if row[col_index] == '':
                        header = headers[col_index]
                        content = get_kinmuchi1() if header == '勤務地' else (get_kaishamei() if header == '会社名' else (get_shokushumei() if header == '職種名' else get_content_by_header(header)))
                        worksheet.update_cell(idx, col_index+1, content)
                        sleep(2)

except Exception as e:
    print(f'例外エラー: {type(e).__name__} \n {str(e)}')

driver.quit()
