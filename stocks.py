#!/usr/bin/env python3

"""
This program gets prices for India stock, India mutual fund, US stock and Crypto prices
via free APIs / web scraping and updates a spreadsheet
Add Holiday handling - http://www.rightline.net/calendar/market-holidays.html
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
import pytest

from bs4 import BeautifulSoup
from datetime import datetime, timedelta

################################################################################
# KEYS
################################################################################
QUANDLKEY = 'ahzB_xyURSjA4V7iHXtT'  # stock
# CURRENCYKEY = '5966705342ec13f9c42e532497e6f060'  # USD rates
# FIXERKEY = '13641e3f360078f91a3294bcc19373d8'  # GBP rates
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
# TARGET = '/Users/vinodverghese/Dropbox/Python/Learning/Completed/Stockapi/stockapi/Financial statement Master.xlsx'

################################################################################
# URLs
################################################################################
GOLDURL = 'https://www.moneymetals.com/precious-metals-charts/gold-price'
SGUTURL = 'https://www.msn.com/en-sg/money/funddetails/fi-F0HKG062P2'
USDURL = 'https://www.exchange-rates.org/currentRates/P/USD'
CADURL = 'https://www.exchange-rates.org/currentRates/P/CAD'
GBPURL = 'https://www.exchange-rates.org/currentRates/P/GBP'
SGDURL = 'https://www.exchange-rates.org/currentRates/P/SGD'

################################################################################
# Stock market holidays
################################################################################
ushols = [
    '2019-1-1', '2020-1-1', '2021-1-1',
    '2019-1-21', '2020-1-20', '2021-1-18', '2022-1-17',
    '2019-2-18', '2020-2-17', '2021-2-15', '2022-2-21',
    '2019-4-19', '2020-4-10', '2021-4-2', '2022-4-15',
    '2019-5-27', '2020-5-25', '2021-5-31', '2022-5-30',
    '2019-7-4', '2020-7-3', '2021-7-5', '2022-7-4',
    '2019-9-2', '2020-9-7', '2021-9-6', '2022-9-5',
    '2019-11-28', '2020-11-26', '2021-11-25', '2022-11-24',
    '2019-12-25', '2020-12-25', '2021-12-24', '2022-12-26'
]

inhols = [
    '2019-3-4', '2019-3-21',
    '2019-4-17',
    '2019-5-1',
    '2019-6-5',
    '2019-8-15',
    '2019-9-2',
    '2019-10-2', '2019-10-28', '2019-10-29',
    '2019-12-25',
    '2020-2-21',
    '2020-3-9',
    '2020-4-14',
    '2020-5-1',
    '2020-6-5',
    '2020-8-15', '2020-8-20',
    '2020-10-2',
    '2020-12-25'
]

sghols = [
    '2019-1-1',
    '2019-2-5', '2019-2-6',
    '2019-4-19',
    '2019-5-1', '2019-5-20',
    '2019-6-5',
    '2019-8-9', '2019-8-12',
    '2019-10-28',
    '2019-12-25'
]

# global vars
logger = None
wb = None
querydte = None


def setup_logger():
    """Logging setup (hierarchy: DIWEC)"""
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    formatter = logging.Formatter('%(levelname)s:%(name)s:%(asctime)s:%(message)s')

    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    file_handler = logging.FileHandler('stocksdebug.log')
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.DEBUG)
    logger.addHandler(file_handler)

    file_handler = logging.FileHandler('stockserr.log')
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.ERROR)
    logger.addHandler(file_handler)

    return logger


def check_file_exists():
    """ check if source file exists """
    if not (os.path.exists(SOURCE)):
        logger.error('%s does not exist' % SOURCE)
        raise SystemExit(99)
    else:
        logger.info('%s exist' % SOURCE)

# def check_output_file_exists():
    # """ check if target file exists """
    # if not (os.path.exists(TARGET)):
    #     logger.error('%s does not exist' % TARGET)
    #     sys.exit(99)
    # else:
    #     logger.info('%s exist' % TARGET)


def backup_input_file():
    """ backup input file """
    # split source path
    path, filename = os.path.split(SOURCE)
    logger.debug('Path : %s / Filename : %s' % (path, filename))

    # split filename
    filewithoutext = os.path.splitext(filename)[0]
    logger.debug('Filename w/o ext : %s' % filewithoutext)

    # split file extension
    ext = os.path.splitext(filename)[1]
    logger.debug('File ext : %s' % ext)

    # Get todays date and format
    now = datetime.now()

    # format date as dd-mm-yyyy
    dte = str(now.day) + '-' + str(now.month) + '-' + str(now.year)
    logger.info('Todays date: %s' % dte)

    # Build up backup path
    bkupfilename = filewithoutext + ' ' + dte + ext
    logger.debug('Backup filename : %s' % bkupfilename)

    bkuppath = path + '/' + bkupfilename
    logger.debug('Backup path : %s' % bkuppath)

    # Backup source file
    logger.info('Backing up..')

    try:
        shutil.copy(SOURCE, BKUPDIR)

    except Exception as e:
        logger.exception('File copy error : %s' % e)
        quit()

    logger.info('Backup complete')

    # Rename file
    oldname = os.path.join(BKUPDIR, filename)
    logger.info('Rename from : %s' % oldname)

    newname = os.path.join(BKUPDIR, bkupfilename)
    logger.info('Rename to : %s' % newname)

    try:
        os.rename(oldname, newname)

    except Exception as e:
        logger.exception('Rename error : %s' % e)
        quit()

    logger.info('Backup renamed')


def get_query_date():
    """ get query date to use """

    # Get todays date and format
    now = datetime.now()

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

    return querydte


def load_excel_workbook():
    """ Load workbook """
    logger.info('Loading workbook')

    try:
        wb = pyxl.load_workbook(SOURCE)

    except Exception as e:
        logger.exception('Workbook load error : %s' % e)
        quit()

    logger.info('Loading complete')

    return wb


def get_forex_rates():
    """ Get forex rates """
    logger.info('Getting forex rates..')

    ws = wb[RATESHEET]

    row = 1

    while (ws.cell(row, 1).value):
        currency = ws.cell(row, 1).value
        logger.info('Currency : %s' % currency)

        if currency[0: 3] == 'USD':
            logger.debug('USD URL : %s' % USDURL)
            page = requests.get(USDURL)
        elif currency[0: 3] == 'GBP':
            logger.debug('GBP URL : %s' % GBPURL)
            page = requests.get(GBPURL)
        elif currency[0: 3] == 'CAD':
            logger.debug('CAD URL : %s' % CADURL)
            page = requests.get(CADURL)
        elif currency[0: 3] == 'SGD':
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
                elif currency == 'USD to SGD':
                    if moretds.find('a', title='Singapore Dollar'):
                        logger.debug('Next sibling: %s' % moretds.next_sibling)
                        usdsgdrate = float(moretds.next_sibling.strong.text)
                        logger.info('USD to SGD Rate : %s' % usdsgdrate)
                        ws.cell(row, 2).value = usdsgdrate
                        row += 1
                        break
                elif currency == 'SGD to INR':
                    if moretds.find('a', title='Indian Rupee'):
                        logger.debug('Next sibling: %s' % moretds.next_sibling)
                        sgdinrrate = float(moretds.next_sibling.strong.text)
                        logger.info('SGD to INR Rate : %s' % sgdinrrate)
                        ws.cell(row, 2).value = sgdinrrate
                        row += 1
                        break
                elif currency == 'CAD to SGD':
                    if moretds.find('a', title='Singapore Dollar'):
                        logger.debug('Next sibling: %s' % moretds.next_sibling)
                        cadsgdrate = float(moretds.next_sibling.strong.text)
                        logger.info('CAD to SGD Rate : %s' % cadsgdrate)
                        ws.cell(row, 2).value = cadsgdrate
                        row += 1
                        break
                elif currency == 'GBP to SGD':
                    if moretds.find('a', title='Singapore Dollar'):
                        logger.debug('Next sibling: %s' % moretds.next_sibling)
                        gbpsgdrate = float(moretds.next_sibling.strong.text)
                        logger.info('GBP to SGD Rate : %s' % gbpsgdrate)
                        ws.cell(row, 2).value = gbpsgdrate
                        row += 1
                        break

            logger.debug('Row : %s' % row)


def get_india_stock_prices():
    """ Get India stock prices """
    logger.info('Getting Indian stock prices..')

    ws = wb[INSHEET]
    # querydte = '2018-11-22'

    # get ticker, call API, update price
    row = 2
    openErr = False
    readErr = False
    parseErr = False

    while (ws.cell(row, 2).value):
        # build API url
        indiaticker = ws.cell(row, 2).value
        logger.info(indiaticker)

        quandlurl = 'https://www.quandl.com/api/v3/datasets/NSE/' + \
            indiaticker + \
            '.json?start_date=' + querydte + '&end_date=' + querydte + \
            '&api_key=' + QUANDLKEY

        logger.debug('URL : %s' % quandlurl)

        # open url
        try:
            f = urllib.request.urlopen(quandlurl)

        except Exception as e:
            logger.exception('URL open error : %s' % e)
            openErr = True

        # read from url
        if not openErr:
            try:
                stockjson = f.read()

            except Exception as e:
                logger.exception('URL read error : %s' % e)
                readErr = True

        # parse json into dictionary
        if not openErr and not readErr:
            try:
                parsed_json = json.loads(stockjson)

            except Exception as e:
                logger.exception('JSON parse error : %s' % e)
                parseErr = True

        logger.debug(json.dumps(parsed_json, indent=4, sort_keys=True))

        # get price from dictionary
        if not openErr and not readErr and not parseErr:
            ws.cell(row, 8).value = float(parsed_json['dataset']['data'][0][1])
            logger.info('Price : %s' % ws.cell(row, 8).value)

        logger.debug('Row : %s' % row)
        row += 1
        time.sleep(0.5)

        openErr = False
        readErr = False
        parseErr = False


def get_india_ut_prices():
    """ Get India UT prices """
    logger.info('Getting Indian UT prices..')

    ws = wb[INUTSHEET]

    # get ticker, call API, update price
    row = 2
    openErr = False
    readErr = False
    parseErr = False

    while (ws.cell(row, 2).value):
        # build API url
        indiautticker = ws.cell(row, 2).value
        logger.info('Ticker : %s' % indiautticker)

        quandlurl = 'https://www.quandl.com/api/v3/datasets/AMFI/' + \
            str(indiautticker) + \
            '.json?start_date=' + querydte + '&end_date=' + querydte + \
            '&api_key=' + QUANDLKEY

        logger.debug('India UT URL : %s' % quandlurl)

        # open url
        try:
            f = urllib.request.urlopen(quandlurl)

        except Exception as e:
            logger.exception('URL open error : %s' % e)
            openErr = True

        # read from url
        if not openErr:
            try:
                stockjson = f.read()

            except Exception as e:
                logger.exception('URL read error : %s' % e)
                readErr = True

        # parse json into dictionary
        if not openErr and not readErr:
            try:
                parsed_json = json.loads(stockjson)
                logger.debug('JSON parser : %s' % parsed_json)

            except Exception as e:
                logger.exception('JSON parser error : %s' % e)
                parseErr = True

        # logger.debug(json.dumps(parsed_json, indent=4, sort_keys=True))
        # get price from dictionary
        if not openErr and not readErr and not parseErr:
            ws.cell(row, 13).value = float(parsed_json['dataset']['data'][0][1])
            logger.info('Price : %s' % ws.cell(row, 13).value)

        logger.debug('Row : %s' % row)
        row += 1
        time.sleep(0.5)

        openErr = False
        readErr = False
        parseErr = False


def get_singapore_stock_prices():
    """ Get Singapore stock prices """
    logger.info('Getting Singapore stock prices..')

    ws = wb[SGSHEET]

    row = 2

    while (ws.cell(row, 2).value):
        sgticker = ws.cell(row, 2).value
        logger.info('SG Ticker : %s' % sgticker)

        if sgticker == 'D05.SI':
            sgtickerurl = 'https://finance.yahoo.com/quote/D05.SI?p=D05.SI&.tsrc=fin-srch'
        elif sgticker == 'O39.SI':
            sgtickerurl = 'https://finance.yahoo.com/quote/O39.SI?p=O39.SI&.tsrc=fin-srch'
        elif sgticker == 'A68U.SI':
            sgtickerurl = 'https://finance.yahoo.com/quote/A68U.SI?p=A68U.SI&.tsrc=fin-srch'
        elif sgticker == 'NR7.SI':
            sgtickerurl = 'https://finance.yahoo.com/quote/NR7.SI?p=NR7.SI&.tsrc=fin-srch'

        # logger.info('SG Stock URL : %s' % sgtickerurl)

        page = requests.get(sgtickerurl)

        # create BS
        soup = BeautifulSoup(page.content, 'lxml')

        divtag = soup.find('div', class_='My(6px) Pos(r) smartphone_Mt(6px)')
        price = divtag.div.span.text

        ws.cell(row, 8).value = float(price)
        logger.info('Price : %s' % ws.cell(row, 8).value)
        row += 1


def get_singapore_ut_prices():
    """ Get Singapore UT price """
    logger.info('Getting SG UT price..')

    ws = wb[SGUTSHEET]

    # load page from url
    try:
        page = requests.get(SGUTURL)

    except Exception as e:
        logger.exception('URL open error : %s' % e)
        quit()

    logger.debug('HTML : %s' % page)

    # create BS
    soup = BeautifulSoup(page.content, 'lxml')
    logger.debug('Soup : %s' % soup.body)

    panel = soup.find('div', {'class': 'precurrentvalue'})
    logger.debug('Tag : %s' % panel)

    price = panel.span.text

    # update price in sheet
    ws['L2'].value = float(price)
    logger.info('UT Price : %s' % ws['L2'].value)


def get_us_stock_prices():
    """ Get US stock prices """
    logger.info('Getting US stock prices..')

    ws = wb[USSHEET]

    row = 2
    openErr = False
    readErr = False
    parseErr = False

    # get ticker, call APIs and update Excel
    while (ws.cell(row, 2).value):
        usticker = ws.cell(row, 2).value
        logger.info('US stock ticker : %s' % usticker)

        iexurl = 'https://api.iextrading.com/1.0/stock/' + usticker + '/book'
        logger.debug('US stock price url : %s' % iexurl)

        # open url
        try:
            f = urllib.request.urlopen(iexurl)

        except Exception as e:
            logger.exception('URL open error : %s' % e)
            openErr = True

        # read url
        if not openErr:
            try:
                stockjson = f.read()

            except Exception as e:
                logger.exception('URL read error : %s' % e)
                readErr = True

        # parse json into dictionary
        if not openErr and not readErr:
            try:
                parsed_json = json.loads(stockjson)
                logger.debug('JSON parser error : %s' % parsed_json)

            except Exception as e:
                logger.exception('JSON parser error : %s' % e)
                parseErr = True

        # logger.debug(json.dumps(parsed_json, indent=4, sort_keys=True))

        # update Excel sheet
        if not openErr and not readErr and not parseErr:
            ws.cell(row, 6).value = float(parsed_json['quote']['latestPrice'])
            logger.info('Price : %s' % ws.cell(row, 6).value)

        logger.debug('Row : %s' % row)
        row += 1
        time.sleep(0.5)

        openErr = False
        readErr = False
        parseErr = False


def get_crypto_prices():
    """ Get Crypto prices """
    logger.info('Getting Crypto prices..')

    ws = wb[CRYPTOSHEET]

    row = 2
    openErr = False
    readErr = False
    parseErr = False

    # get ticker, call APIs and update Excel
    while (ws.cell(row, 2).value):
        cryptoticker = ws.cell(row, 2).value
        logger.info('Crypto ticker : %s' % cryptoticker)

        cryptourl = 'https://min-api.cryptocompare.com/data/price?fsym=' + \
            cryptoticker + \
            '&tsyms=BTC,USD'
        logger.debug('Crypto URL : %s' % cryptourl)

        # open url
        try:
            f = urllib.request.urlopen(cryptourl)

        except Exception as e:
            logger.exception('URL open error : %s' % e)
            openErr = True

        # read url
        if not openErr:
            try:
                stockjson = f.read()

            except Exception as e:
                logger.exception('URL read error : %s' % e)
                readErr = True

        # parse json into dictionary
        if not openErr and not readErr:
            try:
                parsed_json = json.loads(stockjson)
                # login.debug(parsed_json)

            except Exception as e:
                logger.exception('JSON parser error : %s' % e)
                parseErr = True

        # logger.debug(json.dumps(parsed_json, indent=4, sort_keys=True))
        # update Excel sheet
        if not openErr and not readErr and not parseErr:
            ws.cell(row, 5).value = float(parsed_json['USD'])
            logger.info('Price : %s' % ws.cell(row, 5).value)

        logger.debug('Row : %s' % row)
        row += 1
        time.sleep(0.5)

        openErr = False
        readErr = False
        parseErr = False


def get_gold_price():
    # Get Gold price
    logger.info('Getting Gold price..')

    ws = wb[GOLDSHEET]

    # load page from url
    page = requests.get(GOLDURL)
    logger.debug('HTML : %s' % page)

    # create BS
    soup = BeautifulSoup(page.content, 'lxml')
    logger.debug('Soup : %s' % soup.body)

    # find table tag housing price
    panel = soup.find('td', {'class': 'text-center'})
    logger.debug('Tag : %s' % panel)

    # get price and remove non numeric chars
    price = panel.text.replace('$', '').replace(',', '')

    # update price in sheet
    ws['D8'].value = float(price)
    logger.info('Gold price: % s' % ws['D8'].value)


def save_workbook():
    """ save workbook """
    logger.info('Saving workbook..')

    try:
        wb.save(SOURCE)

    except Exception as e:
        logger.exception('Workbook save error : %s' % e)
        quit()

    logger.info('Saved')


if __name__ == '__main__':
    logger = setup_logger()
    check_file_exists()
    backup_input_file()

    querydte = get_query_date()

    wb = load_excel_workbook()

    get_forex_rates()

    get_india_stock_prices()
    get_india_ut_prices()

    get_singapore_stock_prices()
    get_singapore_ut_prices()

    get_us_stock_prices()
    get_crypto_prices()
    get_gold_price()

    save_workbook()
