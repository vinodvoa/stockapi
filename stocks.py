"""
This program gets India stock, India mutual fund, US stock and Crypto prices
via free APIs / web scraping and updates a spreadsheet
Add  SG stocks/ Holiday handling
"""
import sys
import os.path
import urllib.request
import requests
import shutil
import time
import json
import logging
import openpyxl as pyxl

from bs4 import BeautifulSoup
from datetime import datetime, timedelta

################################################################################
# KEYS
################################################################################
QUANDLKEY = 'ahzB_xyURSjA4V7iHXtT'  # stock
CURRENCYKEY = '5966705342ec13f9c42e532497e6f060'  # USD rates
FIXERKEY = '13641e3f360078f91a3294bcc19373d8'  # GBP rates
# ALPHAKEY = 'D9ZW4UF3RPMXif7H'
# MASHAPEKEY = 'OhAQqUguNXmshsiWsbcrJGpk4UI9p1uAoNsjsnjQcjKDPPJvlx'

################################################################################
# Excel sheet name
################################################################################
RATESHEET = 'Rates'
INSHEET = 'IN Stocks'
INUTSHEET = 'IN-UT'
SGSHEET = 'SG Stocks'
SGUTSHEET = 'SG-UT'
USSHEET = 'US Stocks'
CRYPTOSHEET = 'Crypto'
GOLDSHEET = 'Gold'

################################################################################
# Paths
################################################################################
BKUPDIR = '/Volumes/Secomba/vinodverghese/Boxcryptor/Dropbox/Personal/Finance/Backup'
SOURCE = '/Volumes/Secomba/vinodverghese/Boxcryptor/Dropbox/Personal/Finance/Financial statement Master.xlsx'
TARGET = '/Users/vinodverghese/Dropbox/Python/Learning/Completed/Stockapi/stockapi/Financial statement Master.xlsx'

################################################################################
# URLs
################################################################################
GOLDURL = 'https://www.moneymetals.com/precious-metals-charts/gold-price'
SGUTURL = 'https://www.msn.com/en-sg/money/funddetails/fi-F0HKG062P2'
USDURL = 'https://www.exchange-rates.org/currentRates/P/USD'
CADURL = 'https://www.exchange-rates.org/currentRates/P/CAD'
GBPURL = 'https://www.exchange-rates.org/currentRates/P/GBP'
SGDURL = 'https://www.exchange-rates.org/currentRates/P/SGD'

# INMFURL = 'https://www.amfiindia.com/spages/NAVAll.txt'

# logging setup (hierarchy: DIWEC)
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

formatter = logging.Formatter('%(levelname)s:%(name)s:%(asctime)s:%(message)s')

stream_handler = logging.StreamHandler()
stream_handler.setLevel(logging.INFO)
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)

file_handler = logging.FileHandler('debug.log')
file_handler.setFormatter(formatter)
file_handler.setLevel(logging.DEBUG)
logger.addHandler(file_handler)

file_handler = logging.FileHandler('stocks.log')
file_handler.setFormatter(formatter)
file_handler.setLevel(logging.ERROR)
logger.addHandler(file_handler)

################################################################################
# check if source file exists
################################################################################
if not (os.path.exists(SOURCE)):
    logger.error('%s does not exist' % SOURCE)
    sys.exit(99)
else:
    logger.info('%s exist' % SOURCE)

################################################################################
# check if target file exists
################################################################################
if not (os.path.exists(TARGET)):
    logger.error('%s does not exist' % TARGET)
    sys.exit(99)
else:
    logger.info('%s exist' % TARGET)

################################################################################
# split source path
################################################################################
path, filename = os.path.split(SOURCE)
logger.debug('Path : %s / Filename : %s' % (path, filename))

################################################################################
# split filename
################################################################################
filewithoutext = os.path.splitext(filename)[0]
logger.debug('Filename w/o ext : %s' % filewithoutext)

################################################################################
# split file extension
################################################################################
ext = os.path.splitext(filename)[1]
logger.debug('File ext : %s' % ext)

################################################################################
# Get todays date and format
################################################################################
now = datetime.now()

# format date as dd-mm-yyyy
dte = str(now.day) + '-' + str(now.month) + '-' + str(now.year)
logger.info('Todays date: %s' % dte)

################################################################################
# Build up backup path
################################################################################
bkupfilename = filewithoutext + ' ' + dte + ext
logger.debug('Backup filename : %s' % bkupfilename)

bkuppath = path + '/' + bkupfilename
logger.debug('Backup path : %s' % bkuppath)

################################################################################
# Backup source file
################################################################################
logger.info('Backing up..')

try:
    shutil.copy(SOURCE, BKUPDIR)

except Exception as e:
    logger.exception('File copy error : %s' % e)
    quit()

################################################################################
# Rename file
################################################################################
oldname = os.path.join(BKUPDIR, filename)
logger.info('Rename from : %s' % oldname)

newname = os.path.join(BKUPDIR, bkupfilename)
logger.info('Rename to : %s' % newname)

try:
    os.rename(oldname, newname)

except Exception as e:
    logger.exception('Rename error : %s' % e)
    quit()

logger.info('Backup complete')

################################################################################
# Check if weekend
################################################################################
if now.weekday() in range(0, 5):  # Weekdays
    if now.weekday() == 0:  # Mon
        now = datetime.now() - timedelta(3)
        querydte = str(now.year) + '-' + str(now.month) + '-' + str(now.day)
    else:
        querydte = str(now.year) + '-' + str(now.month) + '-' + str(now.day - 1)
else:
    if now.weekday() == 5:  # Sat
        now = datetime.now() - timedelta(1)
        querydte = str(now.year) + '-' + str(now.month) + '-' + str(now.day)
    else:
        if now.weekday() == 6:  # Sun
            now = datetime.now() - timedelta(2)
            querydte = str(now.year) + '-' + str(now.month) + '-' + str(now.day)

logger.debug('Weekday : %s' % now.weekday())
logger.info('Query date : %s' % querydte)

################################################################################
# Load workbook
################################################################################
logger.info('Loading workbook')

try:
    wb = pyxl.load_workbook(SOURCE)

except Exception as e:
    logger.exception('Workbook load error : %s' % e)
    quit()

logger.info('Loading complete')

################################################################################
# Get forex rates
################################################################################
logger.info('Getting forex rates..')

ws = wb[RATESHEET]

row = 1

while (ws.cell(row, 1).value):
    currency = ws.cell(row, 1).value
    logger.info('Currency : %s' % currency)

    if currency[0: 3] == 'USD':
        logger.debug('USD URL : %s' % USDURL)
        page = requests.get(USDURL)
    else:
        if currency[0: 3] == 'GBP':
            logger.debug('GBP URL : %s' % GBPURL)
            page = requests.get(GBPURL)
        else:
            if currency[0: 3] == 'CAD':
                logger.debug('CAD URL : %s' % CADURL)
                page = requests.get(CADURL)
            else:
                if currency[0: 3] == 'SGD':
                    logger.debug('SGD URL : %s' % SGDURL)
                    page = requests.get(SGDURL)

    logger.debug('Forex page : %s' % page)

    # create BS
    soup = BeautifulSoup(page.content, 'lxml')

    tr = soup.findAll('tr')
    logger.debug('Table rows : %s' % tr)

    for td in tr:
        tds = td.findAll('td')
        logger.debug('Table detail: %s' % tds)

        for moretds in tds:
            logger.debug('More table details: %s' % moretds)

            if currency == 'USD to INR':
                if moretds.find('a', title='Indian Rupee'):
                    logger.debug('Next sibling: %s' % moretds.next_sibling)
                    usdinrrate = float(moretds.next_sibling.strong.text)
                    logger.info('USD to INR Rate : %s' % usdinrrate)
                    ws.cell(row, 2).value = usdinrrate
                    row += 1
                    break
            else:
                if currency == 'USD to SGD':
                    if moretds.find('a', title='Singapore Dollar'):
                        logger.debug('Next sibling: %s' % moretds.next_sibling)
                        usdsgdrate = float(moretds.next_sibling.strong.text)
                        logger.info('USD to SGD Rate : %s' % usdsgdrate)
                        ws.cell(row, 2).value = usdsgdrate
                        row += 1
                        break
                else:
                    if currency == 'SGD to INR':
                        if moretds.find('a', title='Indian Rupee'):
                            logger.debug('Next sibling: %s' % moretds.next_sibling)
                            sgdinrrate = float(moretds.next_sibling.strong.text)
                            logger.info('SGD to INR Rate : %s' % sgdinrrate)
                            ws.cell(row, 2).value = sgdinrrate
                            row += 1
                            break
                    else:
                        if currency == 'CAD to SGD':
                            if moretds.find('a', title='Singapore Dollar'):
                                logger.debug('Next sibling: %s' % moretds.next_sibling)
                                cadsgdrate = float(moretds.next_sibling.strong.text)
                                logger.info('CAD to SGD Rate : %s' % cadsgdrate)
                                ws.cell(row, 2).value = cadsgdrate
                                row += 1
                                break
                        else:
                            if currency == 'GBP to SGD':
                                if moretds.find('a', title='Singapore Dollar'):
                                    logger.debug('Next sibling: %s' % moretds.next_sibling)
                                    gbpsgdrate = float(moretds.next_sibling.strong.text)
                                    logger.info('GBP to SGD Rate : %s' % gbpsgdrate)
                                    ws.cell(row, 2).value = gbpsgdrate
                                    row += 1
                                    break

        # logger.info('Rate : %s' % ws.cell(row, 2).value)
        # ws.cell(row, 2).value = rate

        logger.debug('Row : %s' % row)

# row = 1
#
# while (ws.cell(row, 1).value):
#     if ws.cell(row, 1).value == 'CAD to SGD':
#         ws.cell(row, 2).value = cadsgdrate
#     else:
#         if ws.cell(row, 1).value == 'GBP to SGD':
#             ws.cell(row, 2).value = gbpsgdrate
#         else:
#             if ws.cell(row, 1).value == 'SGD to INR':
#                 ws.cell(row, 2).value = sgdinrrate
#             else:
#                 if ws.cell(row, 1).value == 'USD to INR':
#                     ws.cell(row, 2).value = usdinrrate
#                 else:
#                     if ws.cell(row, 1).value == 'USD to SGD':
#                         ws.cell(row, 2).value = usdsgdrate
#
#     logger.debug('Rate: ')

# ################################################################################
# # Get India stock prices
# ################################################################################
# logger.info('Getting Indian stock prices..')
#
# ws = wb[INSHEET]
# # querydte = '2018-11-22'
#
# # get ticker, call API, update price
# row = 2
# openErr = False
# readErr = False
# parseErr = False
#
# while (ws.cell(row, 2).value):
#     # build API url
#     indiaticker = ws.cell(row, 2).value
#     logger.info(indiaticker)
#
#     quandlurl = 'https://www.quandl.com/api/v3/datasets/NSE/' + \
#         indiaticker + \
#         '.json?start_date=' + querydte + '&end_date=' + querydte + \
#         '&api_key=' + QUANDLKEY
#
#     logger.debug('URL : %s' % quandlurl)
#
#     # https: // www.quandl.com / api / v3 / datasets / BSE / ' + \
#     #   indiaticker + \
#     #   '.json?start_date = ' + querydte + ' & end_date = ' + querydte + \
#     #   ' & api_key =' + QUANDLKEY
#
#     # open url
#     try:
#         f = urllib.request.urlopen(quandlurl)
#
#     except Exception as e:
#         logger.exception('URL open error : %s' % e)
#         openErr = True
#
#     # read from url
#     if not openErr:
#         try:
#             stockjson = f.read()
#
#         except Exception as e:
#             logger.exception('URL read error : %s' % e)
#             readErr = True
#
#     # parse json into dictionary
#     if not openErr and not readErr:
#         try:
#             parsed_json = json.loads(stockjson)
#
#         except Exception as e:
#             logger.exception('JSON parse error : %s' % e)
#             parseErr = True
#
#     logger.debug(json.dumps(parsed_json, indent=4, sort_keys=True))
#
#     # get price from dictionary
#     if not openErr and not readErr and not parseErr:
#         ws.cell(row, 8).value = float(parsed_json['dataset']['data'][0][1])
#         logger.info('Price " %s' % ws.cell(row, 8).value)
#
#     logger.debug('Row : %s' % row)
#     row += 1
#     time.sleep(0.5)
#
#     openErr = False
#     readErr = False
#     parseErr = False
#
# ################################################################################
# # Get India UT prices
# ################################################################################
# logger.info('Getting Indian UT prices..')
#
# ws = wb[INUTSHEET]
#
# # get ticker, call API, update price
# row = 2
# openErr = False
# readErr = False
# parseErr = False
#
# while (ws.cell(row, 2).value):
#     # build API url
#     indiautticker = ws.cell(row, 2).value
#     logger.info('Ticker : %s' % indiautticker)
#
#     quandlurl = 'https://www.quandl.com/api/v3/datasets/AMFI/' + \
#         str(indiautticker) + \
#         '.json?start_date=' + querydte + '&end_date=' + querydte + \
#         '&api_key=' + QUANDLKEY
#
#     logger.debug('India UT URL : %s' % quandlurl)
#
#     # open url
#     try:
#         f = urllib.request.urlopen(quandlurl)
#
#     except Exception as e:
#         logger.exception('URL open error : %s' % e)
#         openErr = True
#
#     # read from url
#     if not openErr:
#         try:
#             stockjson = f.read()
#
#         except Exception as e:
#             logger.exception('URL read error : %s' % e)
#             readErr = True
#
#     # parse json into dictionary
#     if not openErr and not readErr:
#         try:
#             parsed_json = json.loads(stockjson)
#             logger.debug('JSON parser : %s' % parsed_json)
#
#         except Exception as e:
#             logger.exception('JSON parser error : %s' % e)
#             parseErr = True
#
#     # logger.debug(json.dumps(parsed_json, indent=4, sort_keys=True))
#     # get price from dictionary
#     if not openErr and not readErr and not parseErr:
#         ws.cell(row, 13).value = float(parsed_json['dataset']['data'][0][1])
#         logger.info('Price : %s' % ws.cell(row, 13).value)
#
#     logger.debug('Row : %s' % row)
#     row += 1
#     time.sleep(0.5)
#
#     openErr = False
#     readErr = False
#     parseErr = False
#
# ################################################################################
# # Get Singapore stock prices
# ################################################################################
# logger.info('Getting Singapore stock prices..')
#
# ws = wb[SGSHEET]
# # querydte = '2018-11-23'
#
# # get ticker, call API, update price
# row = 2
# openErr = False
# readErr = False
# parseErr = False
#
# while (ws.cell(row, 2).value):
#     # build API url
#     sgticker = ws.cell(row, 2).value
#     logger.info('SG ticker : %s' % sgticker)
#
#     quandlurl = 'https://www.quandl.com/api/v3/datasets/XSES/' + \
#         sgticker + \
#         '.json?start_date=' + querydte + '&end_date=' + querydte + \
#         '&api_key=' + QUANDLKEY
#
#     logger.debug('SG stock price url : %s' % quandlurl)
#
#     # open url
#     try:
#         f = urllib.request.urlopen(quandlurl)
#
#     except Exception as e:
#         logger.exception('URL open error : %s' % e)
#         openErr = True
#
#     # read from url
#     if not openErr:
#         try:
#             stockjson = f.read()
#
#         except Exception as e:
#             logger.exception('URL read error : %s' % e)
#             readErr = True
#
#     # parse json into dictionary
#     if not openErr and not readErr:
#         try:
#             parsed_json = json.loads(stockjson)
#
#         except Exception as e:
#             logger.exception('JSON parser error : %s' % e)
#             parseErr = True
#
#     # logger.debug(json.dumps(parsed_json, indent=4, sort_keys=True))
#     # get price from dictionary
#     if not openErr and not readErr and not parseErr:
#         ws.cell(row, 8).value = float(parsed_json['dataset']['data'][0][1])
#         logger.info('Price : %s' % ws.cell(row, 8).value)
#
#     logger.debug('Row : %s' % row)
#     row += 1
#     time.sleep(0.5)
#
#     openErr = False
#     readErr = False
#     parseErr = False
#
# ################################################################################
# # Get US stock prices
# ################################################################################
# logger.info('Getting US stock prices..')
#
# ws = wb[USSHEET]
#
# row = 2
# openErr = False
# readErr = False
# parseErr = False
#
# # get ticker, call APIs and update Excel
# while (ws.cell(row, 2).value):
#     usticker = ws.cell(row, 2).value
#     logger.info('US stock ticker : %s' % usticker)
#
#     iexurl = 'https://api.iextrading.com/1.0/stock/' + usticker + '/book'
#     logger.debug('US stock price url : %s' % iexurl)
#
#     # open url
#     try:
#         f = urllib.request.urlopen(iexurl)
#
#     except Exception as e:
#         logger.exception('URL open error : %s' % e)
#         openErr = True
#
#     # read url
#     if not openErr:
#         try:
#             stockjson = f.read()
#
#         except Exception as e:
#             logger.exception('URL read error : %s' % e)
#             readErr = True
#
#     # parse json into dictionary
#     if not openErr and not readErr:
#         try:
#             parsed_json = json.loads(stockjson)
#             logger.debug('JSON parser error : %s' % parsed_json)
#
#         except Exception as e:
#             logger.exception('JSON parser error : %s' % e)
#             parseErr = True
#
#     # logger.debug(json.dumps(parsed_json, indent=4, sort_keys=True))
#
#     # update Excel sheet
#     if not openErr and not readErr and not parseErr:
#         ws.cell(row, 6).value = float(parsed_json['quote']['latestPrice'])
#         logger.info(ws.cell(row, 6).value)
#
#     row += 1
#     time.sleep(0.5)
#
#     openErr = False
#     readErr = False
#     parseErr = False
#
# ################################################################################
# # Get Crypto prices
# ################################################################################
# logger.info('Getting Crypto prices..')
#
# ws = wb[CRYPTOSHEET]
#
# row = 2
# openErr = False
# readErr = False
# parseErr = False
#
# # get ticker, call APIs and update Excel
# while (ws.cell(row, 2).value):
#     cryptoticker = ws.cell(row, 2).value
#     logger.info('Crypto ticker : %s' % cryptoticker)
#
#     cryptourl = 'https://min-api.cryptocompare.com/data/price?fsym=' + \
#         cryptoticker + \
#         '&tsyms=BTC,USD'
#     logger.debug('Crypto URL : %s' % cryptourl)
#
#     # open url
#     try:
#         f = urllib.request.urlopen(cryptourl)
#
#     except Exception as e:
#         logger.exception('URL open error : %s' % e)
#         openErr = True
#
#     # read url
#     if not openErr:
#         try:
#             stockjson = f.read()
#
#         except Exception as e:
#             logger.exception('URL read error : %s' % e)
#             readErr = True
#
#     # parse json into dictionary
#     if not openErr and not readErr:
#         try:
#             parsed_json = json.loads(stockjson)
#             # login.debug(parsed_json)
#
#         except Exception as e:
#             logger.exception('JSON parser error : %s' % e)
#             parseErr = True
#
#     # logger.debug(json.dumps(parsed_json, indent=4, sort_keys=True))
#     # update Excel sheet
#     if not openErr and not readErr and not parseErr:
#         ws.cell(row, 5).value = float(parsed_json['USD'])
#         logger.info('Price : %s' % ws.cell(row, 5).value)
#
#     logger.debug('Row : %s' % row)
#     row += 1
#     time.sleep(0.5)
#
#     openErr = False
#     readErr = False
#     parseErr = False
#
# ################################################################################
# # Get Gold price
# ################################################################################
# logger.info('Getting Gold price..')
#
# ws = wb[GOLDSHEET]
#
# # load page from url
# page = requests.get(GOLDURL)
# # logger.debug('HTML : %s' % page)
#
# # create BS
# soup = BeautifulSoup(page.content)
# # logger.debug('Soup : %s' % soup.body)
#
# # find table tag housing price
# panel = soup.find('td', {'class': 'text-center'})
# # logger.debug('Tag : %s' % panel)
#
# # get price and remove non numeric chars
# price = panel.text.replace('$', '').replace(',', '')
#
# # update price in sheet
# ws['D8'].value = float(price)
# logger.info('Price: % s' % ws['D8'].value)
#
# ################################################################################
# # Get Singapore UT price
# ################################################################################
# logger.info('Getting SG UT price..')
#
# ws = wb[SGUTSHEET]
#
# # load page from url
# try:
#     page = requests.get(SGUTURL)
#
# except Exception as e:
#     logger.exception('URL open error : %s' % e)
#     quit()
#
# logger.debug('HTML : %s' % page)
#
# # create BS
# soup = BeautifulSoup(page.content)
# logger.debug('Soup : %s' % soup.body)
#
# # find table tag housing price
# panel = soup.find('div', {'class': 'precurrentvalue'})
# logger.debug('Tag : %s' % panel)
#
# # get price
# price = panel.span.text
#
# # update price in sheet
# ws['L2'].value = float(price)
# logger.info('Price : %s' % ws['L2'].value)
#
###############################################################################
# save workbook
###############################################################################
logger.info('Saving workbook..')

try:
    wb.save(SOURCE)

except Exception as e:
    logger.exception('Workbook save error : %s' % e)
    quit()

logger.info('Saved')
