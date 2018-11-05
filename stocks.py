"""
This program gets US and India stock prices via APIs of Alphadantage and Quandl
respectively and updates a spreadsheet
"""
import sys
import os
import os.path
import shutil

import openpyxl as pyxl
import urllib.request
import json
import time

# ALPHAKEY = 'D9ZW4UF3RPMXif7H'
QUANDLKEY = 'ahzB_xyURSjA4V7iHXtT'
USSHEET = 'US Stocks'
INSHEET = 'IN Stocks'
CRYPTOSHEET = 'Crypto'
SOURCE = '/Users/vinodverghese/Dropbox/Personal/Finance/Financial statement Master.xlsm'
TARGET = '/Users/vinodverghese/Dropbox/Python/Financial statement Master.xlsx'

# backup file
if not (os.path.exists(SOURCE)):
    print('%s does not exist' % SOURCE)
    sys.exit(99)

try:
    copyfile(SOURCE, TARGET)

except Exception as e:
    print('Copy Error')
    print(e)
    quit()

# # rename file
# newname =
# os.rename(PATH,)

# load workbook
try:
    wb = pyxl.load_workbook('Financial statement Master.xlsx')

except Exception as e:
    print('1. Workbook Load Error')
    print(e)
    quit()

# Process INDIA spreadsheet
ws = wb[INSHEET]

# get ticker, call API, update price
row = 2
openErr = False
readErr = False
parseErr = False

while (ws.cell(row, 2).value):
    # build API url
    cryptoticker = ws.cell(row, 2).value
    print(indiaticker)

    quandlurl = 'https://www.quandl.com/api/v3/datasets/NSE/' + \
        indiaticker + \
        '.json?start_date=2018-10-26&end_date=2018-10-26&api_key=' + QUANDLKEY
    # print(quandlurl)

    # open url
    try:
        f = urllib.request.urlopen(quandlurl)

    except Exception as e:
        print('1. URL Open Error')
        print(e)
        openErr = True

    # read from url
    if not openErr:
        try:
            stockjson = f.read()

        except Exception as e:
            print('1. URL Read Error')
            print(e)
            readErr = True

    # parse json into dictionary
    if not openErr and not readErr:
        try:
            parsed_json = json.loads(stockjson)

        except Exception as e:
            print('1. Parsing Error')
            print(e)
            parseErr = True

    # print(json.dumps(parsed_json, indent=4, sort_keys=True))
    # get price from dictionary
    if not openErr and not readErr and not parseErr:
        ws.cell(row, 8).value = float(parsed_json['dataset']['data'][0][1])
        print(ws.cell(row, 8).value)

    print(row)
    row += 1
    time.sleep(1)

# get US sheet
ws = wb[USSHEET]

row = 2
openErr = False
readErr = False
parseErr = False

# get ticker, call APIs and update Excel
while (ws.cell(row, 2).value):
    usticker = ws.cell(row, 2).value
    print(usticker)

    iexurl = 'https://api.iextrading.com/1.0/stock/' + usticker + '/book'
    # print(iexurl)

    # open url
    try:
        f = urllib.request.urlopen(iexurl)

    except Exception as e:
        print('2. URL Open Error')
        print(e)
        openErr = True

    # read url
    if not openErr:
        try:
            stockjson = f.read()

        except Exception as e:
            print('2. URL Read Error')
            print(e)
            readErr = True

    # parse json into dictionary
    if not openErr and not readErr:
        try:
            parsed_json = json.loads(stockjson)
            # print(parsed_json)

        except Exception as e:
            print('2. Parse Error')
            print(e)
            parseErr = True

    # print(json.dumps(parsed_json, indent=4, sort_keys=True))
    # update Excel sheet
    if not openErr and not readErr and not parseErr:
        ws.cell(row, 6).value = float(parsed_json['quote']['latestPrice'])
        print(ws.cell(row, 6).value)

    print(row)
    row += 1
    time.sleep(1)

    openErr = False
    readErr = False
    parseErr = False

# get Crypto sheet
ws = wb[CRYPTOSHEET]

row = 2
openErr = False
readErr = False
parseErr = False

# get ticker, call APIs and update Excel
while (ws.cell(row, 2).value):
    cryptoticker = ws.cell(row, 2).value
    print(cryptoticker)

    cryptourl = 'https://min-api.cryptocompare.com/data/price?fsym=' + \
        cryptoticker + \
        '&tsyms=BTC,USD'
    # print(cryptourl)

    # open url
    try:
        f = urllib.request.urlopen(cryptourl)

    except Exception as e:
        print('3. URL Open Error')
        print(e)
        openErr = True

    # read url
    if not openErr:
        try:
            stockjson = f.read()

        except Exception as e:
            print('3. URL Read Error')
            print(e)
            readErr = True

    # parse json into dictionary
    if not openErr and not readErr:
        try:
            parsed_json = json.loads(stockjson)
            print(parsed_json)

        except Exception as e:
            print('3. Parse Error')
            print(e)
            parseErr = True

    # print(json.dumps(parsed_json, indent=4, sort_keys=True))
    # update Excel sheet
    if not openErr and not readErr and not parseErr:
        ws.cell(row, 5).value = float(parsed_json['USD'])
        print(ws.cell(row, 5).value)

    print(row)
    row += 1
    time.sleep(1)

    openErr = False
    readErr = False
    parseErr = False

# save workbook
try:
    print('Save')
    wb.save('Financial statement Master.xlsx')

except Exception as e:
    print('Save Error')
    print(e)
    quit()
