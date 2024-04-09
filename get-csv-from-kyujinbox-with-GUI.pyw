import wx
import wx.grid
import os
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

USERNAME = ''
PASSWORD = ''

LOGIN_URL = 'https://secure.kyujinbox.com/login'

CONFIG_FILE_PATH = './config.ini'

class MyFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="KYUJINBOXヘルパー", size=(800, 600))
        
        panel = wx.Panel(self)

        self.text_edit = wx.TextCtrl(panel, style=wx.TE_MULTILINE)
        
        font = self.text_edit.GetFont()
        font.SetPointSize(13)  # Set font size to 12
        font.SetWeight(wx.FONTWEIGHT_BOLD)  # Set bold
        self.text_edit.SetFont(font)

        self.selected_folder_path, text_value = self.load_configuration()
        self.text_edit.SetValue(text_value)

        self.folder_picker = wx.DirPickerCtrl(panel, message="フォルダーを選んでください。", path=self.selected_folder_path, style=wx.DIRP_DEFAULT_STYLE)
        self.folder_picker.Bind(wx.EVT_DIRPICKER_CHANGED, self.on_folder_pick)

        button = wx.Button(panel, label="実行")
        button.Bind(wx.EVT_BUTTON, self.proceed)
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.text_edit, proportion=1, flag=wx.EXPAND | wx.ALL, border=10)
        sizer.Add(self.folder_picker, flag=wx.EXPAND | wx.ALL, border=10)
        sizer.Add(button, flag=wx.ALIGN_RIGHT | wx.ALL, border=10)
        
        panel.SetSizer(sizer)
        
    def load_configuration(self):
        selected_folder_path = wx.StandardPaths.Get().GetDocumentsDir()
        text_value = ""
        
        if os.path.exists(CONFIG_FILE_PATH):
            with open(CONFIG_FILE_PATH, 'r', encoding='utf-8') as config_file:
                selected_folder_path = config_file.readline().strip()
                text_value = config_file.read()

        return selected_folder_path, text_value

    def save_configuration(self):
        with open(CONFIG_FILE_PATH, 'w', encoding='utf-8') as config_file:
            config_file.write(self.selected_folder_path + '\n')
            config_file.write(self.text_edit.GetValue())
        
    def on_folder_pick(self, event):
        self.selected_folder_path = self.folder_picker.GetPath()

    def proceed(self, event):
        self.text_edit.SetDefaultStyle(wx.TextAttr(wx.NullColour, wx.NullColour))

        text_data = self.text_edit.GetValue()
        self.save_configuration()

        lines = text_data.split('\n')

        # Windowsセレニウム起動

        service = Service(executable_path=r'C:\chromedriver\chromedriver.exe')
        chrome_options = Options()
        chrome_options.add_argument('--lang=ja-JP')
        chrome_options.add_experimental_option('detach', True)
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": self.selected_folder_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        driver = webdriver.Chrome(service=service, options=chrome_options)

        # Windowsセレニウム起動

        driver.get(LOGIN_URL)
        sleep(2)

        try:
            el_username = driver.find_element(By.CSS_SELECTOR, "input#login_email")
            el_username.send_keys(USERNAME)
            sleep(2)

            el_password = driver.find_element(By.CSS_SELECTOR, "input#login_password")
            el_password.send_keys(PASSWORD)
            sleep(2)

            #wx.MessageBox("This is a simple alert!", "Alert", wx.OK | wx.ICON_INFORMATION)

            btn_submit = driver.find_element(By.CSS_SELECTOR, "button#BtnLogin")
            btn_submit.submit()
            sleep(1)

            #print('ユーザー名、パスワードでログイン済')

            driver.maximize_window()

            for line_index, line in enumerate(lines, start=1):

                variables = line.split(',')

                if len(variables) != 5:
                    continue

                driver.get('https://saiyo.kyujinbox.com/ptr/l-accounts')

                company_name, confirm_type, category, start_date, end_date = map(str.strip, variables)

                el_confirm = driver.find_element(By.XPATH, f"//a[contains(@class, 'c-switch__item') and contains(text(), '{confirm_type}')]")
                el_confirm.click();
                
                sleep(1);

                el_company = driver.find_element(By.XPATH, f"//a[contains(@class, 'c-link') and contains(text(), '{company_name}')]")
                el_company.click();

                if confirm_type == '直接投稿':
                    el_adredirect = driver.find_element(By.XPATH, "//a[contains(@class, 'Ad') and contains(text(), '有料利用状況')]")
                    el_adredirect.click();
                
                    el_reporting = driver.find_element(By.XPATH, "//a[contains(@class, 'c-tab__link') and contains(text(), 'レポーティング')]")
                    el_reporting.click();
                
                    el_category = driver.find_element(By.XPATH, f"//p[contains(@class, 'TabBtn') and contains(text(), '{category}')]")
                    el_category.click();
                
                    el_dateselect = driver.find_element(By.CSS_SELECTOR, ".ReportFilterForm > .DateModal-select")
                    el_dateselect.click();

                    ip_from = driver.find_element(By.CSS_SELECTOR, "input[name='modal_from']")
                    sleep(1)
                    ip_from.send_keys(Keys.CONTROL, 'a', Keys.DELETE)
                    ip_from.send_keys(start_date)
                    ip_from.send_keys(Keys.TAB)
                    ip_to = driver.find_element(By.CSS_SELECTOR, "input[name='modal_to']")
                    sleep(1)
                    ip_to.send_keys(Keys.CONTROL, 'a', Keys.DELETE)
                    ip_to.send_keys(end_date)
                    ip_from.send_keys(Keys.TAB)
                    btn_submit = driver.find_element(By.CSS_SELECTOR, ".DateModal_footer > button.BtnBlue")
                    sleep(2)
                    btn_submit.click()

                elif confirm_type == 'クローリング・フィード':
                    el_report = driver.find_element(By.XPATH, "//li/a[contains(text(), 'レポート')]")
                    el_report.click();
                
                    el_category = driver.find_element(By.XPATH, f"//ul/li[contains(text(), '{category}')]")
                    el_category.click();
                
                    el_dateselect = driver.find_element(By.CSS_SELECTOR, ".displayPeriod > .modal-date-select")
                    el_dateselect.click();

                    ip_from = driver.find_element(By.CSS_SELECTOR, "input[name='modal_from']")
                    sleep(1)
                    ip_from.send_keys(Keys.CONTROL, 'a', Keys.DELETE)
                    ip_from.send_keys(start_date)
                    ip_from.send_keys(Keys.TAB)
                    ip_to = driver.find_element(By.CSS_SELECTOR, "input[name='modal_to']")
                    sleep(1)
                    ip_to.send_keys(Keys.CONTROL, 'a', Keys.DELETE)
                    ip_to.send_keys(end_date)
                    ip_to.send_keys(Keys.TAB)
                    btn_submit = driver.find_element(By.CSS_SELECTOR, "button.js-apply-btn")
                    sleep(2)
                    btn_submit.click()
                
                sleep(1)

                el_download = driver.find_element(By.CSS_SELECTOR, "a.js-csv-download")
                el_download.click()

            wx.MessageBox(f'処理完了～', 'メッセージー', wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f'{type(e).__name__} \n {str(e)}', '例外エラー', wx.OK | wx.ICON_ASTERISK)

app = wx.App()
frame = MyFrame()
frame.Show()
app.MainLoop()