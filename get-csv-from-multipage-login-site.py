from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

USERNAME = ''
PASSWORD = ''

LOGIN_URL = 'https://agent-bank.com/service/login?next=%2Fhome'
TARGET_URL = 'https://agent-bank.com/service/job/list?...'

OUTPUT_FILE = 'urls.csv'

# Windowsセレニウム起動

service = Service(executable_path=r'D:\chromedriver\chromedriver.exe')
chrome_options = Options()
chrome_options.add_experimental_option('detach', True)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Windowsセレニウム起動

driver.get(LOGIN_URL)
sleep(2)

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

    driver.get(TARGET_URL)
    sleep(3)

    number_of_jobs = driver.find_element(By.CSS_SELECTOR, ".count > .all").text;
    print(f'求人検索結果一覧: {number_of_jobs}件');

    with open(OUTPUT_FILE, 'a') as file:
        while True:       
            current_page = driver.find_element(By.CSS_SELECTOR, "li.number.active").text
            total_pages = driver.find_element(By.CSS_SELECTOR, "li.number:last-child").text
            print(f'URL取得中: {current_page}/{total_pages}')

            ell_job_links = driver.find_elements(By.CSS_SELECTOR, "div.leading-none > a.title")

            for el_job_link in ell_job_links:
                href = el_job_link.get_attribute("href")
                file.write(href + '\n')
                #print(href)

            # 次のページがある場合
            el_next_button = driver.find_element(By.CSS_SELECTOR, "button.btn-next")
            
            # 最後のページで、習得完了
            if 'disabled' in el_next_button.get_attribute("class"):
                break

            # 次ボタンを押下し、次のページのロードで待機
            el_next_button.click()
            sleep(3)

except Exception as e:
    print(f'例外エラー: {type(e).__name__} \n {str(e)}')

driver.quit()