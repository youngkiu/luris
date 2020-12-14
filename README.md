# luris
Land Use Regulations Information Service http://luris.molit.go.kr/

## Prerequisites

Download the chromedriver for the version of Chrome browser installed on your PC.

https://chromedriver.chromium.org/downloads

## Usage
```
python luris.py -d 경상북도 -s 안동시 -i "표본목록(안동).xls"
```

## Built
```
pyinstaller --onefile luris.py
```
