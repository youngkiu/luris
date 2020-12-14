#-*- coding: utf-8 -*-

import os
import sys
import time
import json
import glob
import math
import datetime
import argparse
import openpyxl
import xlrd
import traceback
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select


def __wait_for_time(future, now, display_period=3):
    print('The program will start at %s' % future.strftime("%Y-%m-%d %H:%M:%S"))
    left = future - now
    if left.days >= 0 and left.seconds > 0:
        # https://stackoverflow.com/questions/3160699/python-progress-bar
        print('The progress of the waiting time')
        toolbar_width = math.ceil(left.seconds / display_period)
        sleep_period = display_period
        start = now
        for i in tqdm(range(toolbar_width)):
            time.sleep(sleep_period)

            predict = start + datetime.timedelta(0, (i + 1) * display_period)
            now = datetime.datetime.now()
            if predict < now:
                sleep_period -= 0.0001
                # print('[Debug] decrease sleep_period(%f) - predict:%s, now:%s' % (
                #     sleep_period, predict.strftime('%H:%M:%S'), now.strftime('%H:%M:%S')))
            else:
                sleep_period += 0.001
                # print('[Debug] increase sleep_period(%f) - predict:%s, now:%s' % (
                #     sleep_period, predict.strftime('%H:%M:%S'), now.strftime('%H:%M:%S')))

    now = datetime.datetime.now()
    print('It is now %s' % now.strftime("%Y-%m-%d %H:%M:%S"))


def __parse_umd_ri_bn(umd_ri, gbn_bobn_bubn):
    umd_ri_list = umd_ri.split()
    umd = umd_ri_list[0]
    if len(umd_ri_list) > 1:
        assert len(umd_ri_list) == 2, '[Error] %s' % umd_ri
        ri = umd_ri_list[1]
    else:
        ri = ''

    gbn_idx = gbn_bobn_bubn.find('산')
    if gbn_idx < 0:
        gbn = '일반'
        bobn_start_idx = 0
    else:
        assert gbn_idx == 0, '[Error] %s' % gbn_bobn_bubn
        gbn = '산'
        bobn_start_idx = 1

    hyphen_idx = gbn_bobn_bubn.rfind('-')
    if hyphen_idx < 0:
        bobn = gbn_bobn_bubn[bobn_start_idx:]
        bubn = ''
    else:
        bobn = gbn_bobn_bubn[bobn_start_idx:hyphen_idx]
        bubn = gbn_bobn_bubn[hyphen_idx + 1:]

    print(umd_ri, gbn_bobn_bubn, '-->', umd, ri, gbn, bobn, bubn)

    return umd, ri, gbn, bobn, bubn


def __get_sample_list(xls_file_path):
    sample_list = []

    file_name, extension = os.path.splitext(xls_file_path)
    if extension == '.xlsx':
        wb = openpyxl.load_workbook(xls_file_path)
        ws = wb.active

        for r in ws.rows:
            serial_num = r[2].value
            umd_ri = r[4].value
            gbn_bobn_bubn = r[5].value

            if serial_num and serial_num.isdigit():
                umd, ri, gbn, bobn, bubn = __parse_umd_ri_bn(umd_ri, gbn_bobn_bubn)
                sample_list.append([serial_num, umd, ri, gbn, bobn, bubn])

        wb.close()
    elif extension == '.xls':
        wb = xlrd.open_workbook(xls_file_path)
        ws = wb.sheet_by_index(0)

        for i in range(ws.nrows):
            serial_num = ws.row_values(i)[2]
            umd_ri = ws.row_values(i)[4]
            gbn_bobn_bubn = ws.row_values(i)[5]

            if serial_num and serial_num.isdigit():
                umd, ri, gbn, bobn, bubn = __parse_umd_ri_bn(umd_ri, gbn_bobn_bubn)
                sample_list.append([serial_num, umd, ri, gbn, bobn, bubn])

        wb.release_resources()
    else:
        return sample_list

    return sample_list


def __query_and_save_pdf(driver, sido, sgg, umd, ri, gbn, bobn, bubn, serial_num, download_dir):
    # https://dejavuqa.tistory.com/110
    wait = WebDriverWait(driver, 10)
    element = wait.until(EC.presence_of_element_located((By.NAME, "selSido")))
    Select(element).select_by_visible_text(sido)

    driver.implicitly_wait(10)
    element = wait.until(EC.presence_of_element_located((By.NAME, "selSgg")))
    Select(element).select_by_visible_text(sgg)

    driver.implicitly_wait(10)
    element = wait.until(EC.presence_of_element_located((By.NAME, "selUmd")))
    Select(element).select_by_visible_text(umd)

    if ri:
        driver.implicitly_wait(10)
        element = wait.until(EC.presence_of_element_located((By.NAME, "selRi")))
        Select(element).select_by_visible_text(ri)

    element = wait.until(EC.presence_of_element_located((By.NAME, "landGbn")))
    Select(element).select_by_visible_text(gbn)

    assert bobn.isdigit(), '[Error] not digit (%s)' % bobn
    wait.until(EC.presence_of_element_located((By.NAME, "bobn"))).send_keys(bobn)

    if bubn:
        assert bubn.isdigit(), '[Error] not digit (%s)' % bubn
        wait.until(EC.presence_of_element_located((By.NAME, "bubn"))).send_keys(bubn)

    wait.until(EC.presence_of_element_located((By.XPATH, '//button[text()="열람"]'))).click()
    driver.implicitly_wait(1000)

    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@class='printa']"))).click()
    driver.implicitly_wait(10)

    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@class='print_bt']"))).click()
    driver.implicitly_wait(100)

    # https://stackoverflow.com/questions/10629815/how-to-switch-to-new-window-in-selenium-for-python
    driver.switch_to.window(driver.window_handles[1])
    driver.implicitly_wait(1000)

    time.sleep(3)
    driver.execute_script('window.print();')
    driver.implicitly_wait(1000)
    time.sleep(1)

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

    assert os.stat(target_file_path).st_size > 0, '[Error] %s file is 0 size' % target_file_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Save the query results of http://luris.molit.go.kr/ as a pdf.', add_help=False)
    parser.add_argument('-d', '--sido', required=True, type=str, help='광역시 및 도')
    parser.add_argument('-s', '--sgg', required=True, type=str, help='시군구')
    parser.add_argument('-i', '--excel', required=True, type=str, help='excel file name')
    parser.add_argument('-f', '--datetime', type=lambda s: datetime.datetime.strptime(s, '%Y-%m-%d %H:%M:%S'), help='date time')
    parser.add_argument('-t', '--time', type=lambda s: datetime.datetime.strptime(s, '%H%M'), help='reservation time')
    # https://www.reddit.com/r/learnpython/comments/3gbkin/overriding_h_short_option_in_argparse/
    parser.add_argument('-h', '--hour', type=int, help='delay hour')
    parser.add_argument('--help', action='help', default=argparse.SUPPRESS, help=argparse._('show this help message and exit'))
    args = parser.parse_args()

    _sido = args.sido
    _sgg = args.sgg
    _xls_file_path = args.excel

    num_of_time_arg = sum(x is not None for x in [args.time, args.hour, args.datetime])
    if num_of_time_arg > 1:
        print('-t(--time) and -h(--hour) and -f(--datetime) options are mutually exclusive')
        sys.exit()

    if num_of_time_arg == 1:
        _now = datetime.datetime.now()
        if args.datetime:
            _future = args.datetime
        elif args.time:
            _future = args.time.replace(year=_now.year, month=_now.month, day=_now.day)
            if _future.time() < _now.time():
                _future += datetime.timedelta(days=1)
        elif args.hour:
            _future = _now + datetime.timedelta(days=args.hour//24, hours=args.hour%24)
        else:
            assert False

        __wait_for_time(_future, _now)

    _sample_list = __get_sample_list(_xls_file_path)
    if not _sample_list:
        print('Incompatible Excel file(%s)' % _xls_file_path)
        sys.exit()

    # https://stackoverflow.com/questions/56897041/how-to-save-opened-page-as-pdf-in-selenium-python
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
    _download_dir = os.path.join(_current_dir, '%s_%s' % (_sido, _sgg))
    if not os.path.exists(_download_dir):
        os.mkdir(_download_dir)

    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
             'savefile.default_directory': _download_dir}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')
    CHROMEDRIVER_PATH = 'chromedriver'

    f = open('error_address_%s_%s.txt' % (_sido, _sgg), 'w')

    num_of_sample = len(_sample_list)
    for i, [_serial_num, _umd, _ri, _gbn, _bobn, _bubn] in enumerate(_sample_list):
        _driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=CHROMEDRIVER_PATH)

        _url = 'http://luris.molit.go.kr/'
        _driver.get(_url)

        address_str = '%s: %s %s %s %s' % (_serial_num, _sido, _sgg, _umd, _ri)
        if _gbn == '산':
            address_str += ' %s' % _gbn
        address_str += ' %s' % _bobn
        if _bubn:
            address_str += '-%s' % _bubn

        print('%4d/%4d, %s' % (i+1, num_of_sample, address_str))

        try:
            __query_and_save_pdf(_driver, _sido, _sgg, _umd, _ri, _gbn, _bobn, _bubn, _serial_num, _download_dir)
        except:
            traceback.print_exc()
            print('[Error] not found - %s', address_str)
            f.write('%s\n' % address_str)
            f.flush()
        finally:
            _driver.quit()

    f.close()
