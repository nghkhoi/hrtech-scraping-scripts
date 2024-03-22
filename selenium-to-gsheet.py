from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import gspread
from oauth2client.service_account import ServiceAccountCredentials

USERNAME = ''
PASSWORD = ''

LOGIN_URL = 'https://agent-bank.com/service/login?next=%2Fhome'

GSHEET_ID = '1A...'
SERVICEACCOUNT_JSON = './file.json'

# Windowsセレニウム起動

service = Service(executable_path=r'D:\chromedriver\chromedriver.exe')
chrome_options = Options()
chrome_options.add_experimental_option('detach', True)
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

try:
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

    for idx, row in enumerate(rows, start=1):
        if row[0] and not row[1]:
            target_url = row[0]
            driver.get(target_url)
            print (f'対応中: {target_url}')
            sleep(2)

            worksheet.update_cell(idx, 2, get_content_by_header('必須要件'))
            worksheet.update_cell(idx, 3, get_content_by_header('仕事内容'))
            worksheet.update_cell(idx, 4, get_content_by_header('仕事の醍醐味'))
            worksheet.update_cell(idx, 5, get_content_by_header('歓迎要件（活躍できる経験）'))
            worksheet.update_cell(idx, 6, get_content_by_header('募集職種'))
            worksheet.update_cell(idx, 7, get_content_by_header('想定年収'))
            sleep(1)
            worksheet.update_cell(idx, 8, get_content_by_header('予定募集人数'))
            worksheet.update_cell(idx, 9, get_content_by_header('管理監督者求人'))
            worksheet.update_cell(idx, 10, get_content_by_header('労働時間区分'))
            sleep(1)
            worksheet.update_cell(idx, 11, get_content_by_header('雇用形態'))
            worksheet.update_cell(idx, 12, get_content_by_header('勤務時間'))
            worksheet.update_cell(idx, 13, get_content_by_header('残業時間'))
            worksheet.update_cell(idx, 14, get_content_by_header('時短勤務'))
            sleep(1)
            worksheet.update_cell(idx, 15, get_content_by_header('選考詳細'))
            worksheet.update_cell(idx, 16, get_kinmuchi1())
            worksheet.update_cell(idx, 17, get_content_by_header('試用期間'))
            worksheet.update_cell(idx, 18, get_content_by_header('試用期間詳細'))
            worksheet.update_cell(idx, 19, get_content_by_header('給与・待遇'))
            sleep(1)
            worksheet.update_cell(idx, 20, get_content_by_header('年間休日'))
            worksheet.update_cell(idx, 21, get_content_by_header('休日・休暇'))
            worksheet.update_cell(idx, 22, get_content_by_header('福利厚生'))
            worksheet.update_cell(idx, 23, get_content_by_header('受動喫煙対策'))
            worksheet.update_cell(idx, 24, get_content_by_header('受動喫煙対策（詳細）'))
            sleep(2)

except Exception as e:
    print(f'例外エラー: {type(e).__name__} \n {str(e)}')

driver.quit()