import os
import time
import json
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from openpyxl import load_workbook


def __get_sample_list(xls_file_path, sheet_name):
    wb = load_workbook(xls_file_path)
    ws = wb.get_sheet_by_name(sheet_name)

    sample_list = []

    for r in ws.rows:
        serial_num = r[2].value
        if serial_num and serial_num.isdigit():
            umd_ri = r[4].value
            gbn_bobn_bubn = r[5].value

            umd_ri_list = umd_ri.split()
            assert len(umd_ri_list) == 2, '[Error] %s' % umd_ri
            umd, ri = umd_ri_list
            gbn_idx = gbn_bobn_bubn.find('산')
            if gbn_idx < 0:
                gbn = '일반'
                bobn_start_idx = 0
            else:
                assert gbn_idx == 0, '[Error] %s' % gbn_bobn_bubn
                gbn = '산'
                bobn_start_idx = 1

            hyphen_idx = gbn_bobn_bubn.find('-')
            if hyphen_idx < 0:
                bobn = gbn_bobn_bubn[bobn_start_idx:]
                bubn = ''
            else:
                bobn = gbn_bobn_bubn[bobn_start_idx:hyphen_idx]
                bubn = gbn_bobn_bubn[hyphen_idx+1:]

            print(serial_num, umd_ri, gbn_bobn_bubn, '-->', umd, ri, gbn, bobn, bubn)

            sample_list.append([umd, ri, gbn, bobn, bubn, serial_num])

    wb.close()

    return sample_list


def __query_and_save_pdf(driver, sido, sgg, umd, ri, gbn, bobn, bubn, serial_num, download_dir):
    print(sido, sgg, umd, ri, gbn, bobn, bubn, serial_num)

    Select(driver.find_element(By.NAME, 'selSido')).select_by_visible_text(sido)
    driver.implicitly_wait(10)

    Select(driver.find_element(By.NAME, 'selSgg')).select_by_visible_text(sgg)
    driver.implicitly_wait(10)

    Select(driver.find_element(By.NAME, 'selUmd')).select_by_visible_text(umd)
    driver.implicitly_wait(10)

    Select(driver.find_element(By.NAME, 'selRi')).select_by_visible_text(ri)
    driver.implicitly_wait(10)

    Select(driver.find_element(By.NAME, 'landGbn')).select_by_visible_text(gbn)
    driver.implicitly_wait(10)

    driver.find_element(By.NAME, 'bobn').send_keys(bobn)
    driver.implicitly_wait(10)
    driver.find_element(By.NAME, 'bubn').send_keys(bubn)
    driver.implicitly_wait(10)

    driver.find_element(By.XPATH, '//button[text()="열람"]').click()
    driver.implicitly_wait(100)

    driver.find_element(By.CLASS_NAME, 'printa').click()
    driver.implicitly_wait(10)
    driver.find_element(By.CLASS_NAME, 'print_bt').click()
    driver.implicitly_wait(100)

    driver.switch_to.window(driver.window_handles[1])
    driver.implicitly_wait(10)

    driver.execute_script('window.print();')
    driver.implicitly_wait(1000)

    target_file_path = os.path.join(download_dir, '%s.pdf' % serial_num)
    if os.path.exists(target_file_path):
        os.remove(target_file_path)

    while not glob.glob(os.path.join(download_dir, '토지이용계획 - LURIS 토지이용규제정보서비스*.pdf')):
        time.sleep(1)

    default_save_name_list = glob.glob(os.path.join(download_dir, '토지이용계획 - LURIS 토지이용규제정보서비스*.pdf'))
    for i, duplicated_file in enumerate(default_save_name_list):
        if i == 0:
            os.rename(os.path.join(download_dir, duplicated_file), target_file_path)
        else:
            os.remove(os.path.join(download_dir, duplicated_file))


if __name__ == "__main__":
    _xls_file_path = '표본목록.xlsx'
    _sheet_name = '표본목록'

    _sample_list = __get_sample_list(_xls_file_path, _sheet_name)

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

    _current_dir = os.path.dirname(os.path.realpath(__file__))
    _download_dir = os.path.join(_current_dir, 'pdf')
    if not os.path.exists(_download_dir):
        os.mkdir(_download_dir)

    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
             'savefile.default_directory': _download_dir}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')
    CHROMEDRIVER_PATH = 'chromedriver'

    _sido = '경상북도'
    _sgg = '안동시'
    for _umd, _ri, _gbn, _bobn, _bubn, _serial_num in _sample_list:
        _driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=CHROMEDRIVER_PATH)

        url = 'http://luris.molit.go.kr/web/index.jsp'
        _driver.get(url)

        try:
            __query_and_save_pdf(_driver, _sido, _sgg, _umd, _ri, _gbn, _bobn, _bubn, _serial_num, _download_dir)
        except:
            print('[Error] not found:', _sido, _sgg, _umd, _ri, _gbn, _bobn, _bubn)
        finally:
            _driver.quit()
