import os
import json
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select


chrome_options = webdriver.ChromeOptions()
settings = {
       "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }

current_dir = os.path.dirname(os.path.realpath(__file__))
download_dir = os.path.join(current_dir, 'pdf')
if not os.path.exists(download_dir):
    os.mkdir(download_dir)

prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings), 'savefile.default_directory': download_dir}
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument('--kiosk-printing')
CHROMEDRIVER_PATH = 'chromedriver'
driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=CHROMEDRIVER_PATH)

url = 'http://luris.molit.go.kr/web/index.jsp'
driver.get(url)

Select(driver.find_element_by_name('selSido')).select_by_visible_text('경상북도')
driver.implicitly_wait(10)

Select(driver.find_element_by_name('selSgg')).select_by_visible_text('안동시')
driver.implicitly_wait(10)

Select(driver.find_element_by_name('selUmd')).select_by_visible_text('와룡면')
driver.implicitly_wait(10)

Select(driver.find_element_by_name('selRi')).select_by_visible_text('중가구리')
driver.implicitly_wait(10)

Select(driver.find_element_by_name('landGbn')).select_by_visible_text('일반')
driver.implicitly_wait(10)

driver.find_element_by_name('bobn').send_keys('63')
driver.implicitly_wait(10)
driver.find_element_by_name('bubn').send_keys('2')
driver.implicitly_wait(10)

driver.find_element(By.XPATH, '//button[text()="열람"]').click()
driver.implicitly_wait(100)

driver.find_element_by_class_name('printa').click()
driver.implicitly_wait(10)
driver.find_element_by_class_name('print_bt').click()
driver.implicitly_wait(100)

driver.switch_to.window(driver.window_handles[1])
driver.implicitly_wait(10)

driver.execute_script('window.print();')
driver.implicitly_wait(100)


target_file_path = os.path.join(download_dir, '1538.pdf')
if os.path.exists(target_file_path):
    os.remove(target_file_path)

os.rename(os.path.join(download_dir, '토지이용계획 - LURIS 토지이용규제정보서비스.pdf'), target_file_path)

for duplicated_file in glob.glob(os.path.join(download_dir, '토지이용계획 - LURIS 토지이용규제정보서비스*.pdf')):
    os.remove(os.path.join(download_dir, duplicated_file))

driver.quit()