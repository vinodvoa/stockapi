"""
This program gets India, US and Crypto prices via APIs and updates a spreadsheet
"""
import sys
import os
import os.path
import shutil

import openpyxl as pyxl
import urllib.request
import json
import time
import datetime

from bs4 import BeautifulSoup

# ALPHAKEY = 'D9ZW4UF3RPMXif7H'
QUANDLKEY = 'ahzB_xyURSjA4V7iHXtT'
USSHEET = 'US Stocks'
INSHEET = 'IN Stocks'
CRYPTOSHEET = 'Crypto'
GOLDSHEET = 'Gold'
SOURCE = '/Volumes/Secomba/vinodverghese/Boxcryptor/Dropbox/Personal/Finance/Financial statement Master.xlsx'
BKUPDIR = '/Volumes/Secomba/vinodverghese/Boxcryptor/Dropbox/Personal/Finance/Backup'
TARGET = '/Users/vinodverghese/Dropbox/Python/Learning/Completed/Stockapi/stockapi/Financial statement Master.xlsx'
URL = 'https://www.moneymetals.com/precious-metals-charts/gold-price'

# check if source file exists
if not (os.path.exists(SOURCE)):
    print('%s does not exist' % SOURCE)
    sys.exit(99)
# else:
#     print('File exists')

# check if target file exists
if not (os.path.exists(TARGET)):
    print('%s does not exist' % TARGET)
    sys.exit(99)
else:
    print('File exists')

# split source path
path, filename = os.path.split(SOURCE)
# print(path)
# print(filename)

# split filename
filewithoutext = os.path.splitext(filename)[0]
# print(filewithoutext)

# split file extension
ext = os.path.splitext(filename)[1]
# print(ext)

# Get todays date and time
now = datetime.datetime.now()

# format date as dd-mm-yyyy
dte = str(now.day) + '-' + str(now.month) + '-' + str(now.year)

# build up backup path
bkupfilename = filewithoutext + ' ' + dte + ext
bkuppath = path + '/' + bkupfilename
# print(bkuppath)

# backup file
print('Backing up..')
try:
    shutil.copy(SOURCE, BKUPDIR)

except Exception as e:
    print('Copy Error')
    print(e)
    quit()

# rename file
oldname = os.path.join(BKUPDIR, filename)
# print(oldname)

newname = os.path.join(BKUPDIR, bkupfilename)
# print(newname)

try:
    os.rename(oldname, newname)

except Exception as e:
    print('File Rename Error')
    print(e)
    quit()

print('Backup complete')

# load workbook
print('Loading workbook')

try:
    wb = pyxl.load_workbook(SOURCE)

except Exception as e:
    print('1. Workbook Load Error')
    print(e)
    quit()

print('Loading complete')

# Process INDIA spreadsheet
print('Getting Indian stock prices..')

ws = wb[INSHEET]

# get ticker, call API, update price
row = 2
openErr = False
readErr = False
parseErr = False

while (ws.cell(row, 2).value):
    # build API url
    indiaticker = ws.cell(row, 2).value
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

    # print(row)
    row += 1
    time.sleep(0.5)

# Process US sheet
print('Getting US stock prices..')

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
    time.sleep(0.5)

    openErr = False
    readErr = False
    parseErr = False

# process Crypto sheet
print('Getting Crypto prices..')

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
            # print(parsed_json)

        except Exception as e:
            print('3. Parse Error')
            print(e)
            parseErr = True

    # print(json.dumps(parsed_json, indent=4, sort_keys=True))
    # update Excel sheet
    if not openErr and not readErr and not parseErr:
        ws.cell(row, 5).value = float(parsed_json['USD'])
        print(ws.cell(row, 5).value)

    # print(row)
    row += 1
    time.sleep(0.5)

    openErr = False
    readErr = False
    parseErr = False

# process Gold sheet
print('Getting Gold price..')

ws = wb[GOLDSHEET]

# load page from url
page = requests.get(URL)

# create BS
soup = BeautifulSoup(page.content)
# print(soup.body)

# find table tag housing price
panel = soup.find('td', {'class': 'text-center'})
# print(panel)

# get price and remove non numeric chars
price = panel.text.replace('$', '').replace(',', '')
print(price)

# update price in sheet
ws['D8'] = float(price)

# save workbook
print('Saving workbook..')

try:
    wb.save(SOURCE)

except Exception as e:
    print('Save Error')
    print(e)
    quit()

print('Saved')
