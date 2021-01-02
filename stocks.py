#!/Users/vinodverghese/anaconda3/bin/ python3

"""
This program gets prices for India stock, India mutual fund, US stock & Crypto
prices via free APIs / web scraping and updates a spreadsheet
Add Holiday handling - http://www.rightline.net/calendar/market-holidays.html
"""
import os
import os.path
import requests
import shutil
import time
import logging
import openpyxl as pyxl
import datetime

from bs4 import BeautifulSoup

###############################################################################
# Excel sheet names
###############################################################################
RATESHEET = 'FX Rates'
INSHEET = 'IN Stocks'
INWLSHEET = 'IN WL'
INUTSHEET = 'IN UT'
SGSHEET = 'SG Stocks'
SGUTSHEET = 'SG UT'
USSHEET = 'US Stocks'
USWLSHEET = 'US WL'
CRYPTOSHEET = 'Crypto'
GOLDSHEET = 'Gold'
FSSHEET = 'FS'
EODSHEET = 'EOD'

###############################################################################
# Paths
###############################################################################
BKUPDIR = '/Volumes/Secomba/vinodverghese/Boxcryptor/Dropbox/Personal/Finance/Backup'
SOURCE = '/Volumes/Secomba/vinodverghese/Boxcryptor/Dropbox/Personal/Finance/Financial statement Master.xlsx'


###############################################################################
# URLs
###############################################################################
SGUTURL = 'https://www.ocbc.com/rates/daily_price_unit_trust.html'
USDFXURL = 'https://www.exchange-rates.org/currentRates/P/USD'
CADFXURL = 'https://www.exchange-rates.org/currentRates/P/CAD'
GBPFXURL = 'https://www.exchange-rates.org/currentRates/P/GBP'
SGDFXURL = 'https://www.exchange-rates.org/currentRates/P/SGD'
STATSURL = 'https://finance.yahoo.com/quote/AAPL/key-statistics?p=AAPL'

# global vars
DELAY = 4
logger = None
wb = None
EXCLUDE = ['IVV', 'VFH', 'MCHI', 'VHT', 'SLYG', 'XLE', 'DVY', 'VOE']


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


def check_urls(url):
    page = requests.get(url)
    return page.status_code


def check_file_exists(filename):
    """ check if source file exists """
    if not (os.path.exists(filename)):
        logger.error('%s does not exist' % filename)
        raise SystemExit(99)
    else:
        logger.info('%s exist' % filename)
        return True


def get_date(format):
    # Get todays date
    now = datetime.datetime.now()

    # format date as dd-mm-yyyy
    if format == 1:
        dte = str(now.day) + '-' + str(now.month) + '-' + str(now.year)

    # format date as dd-mm-yyyy hh:mm:ssss
    if format == 2:
        dte = str(now.day) + '-' + str(now.month) + '-' + str(now.year) + ' ' + \
              str(now.hour) + ':' + str(now.minute) + ':' + str(now.second)

    logger.info('Todays date: %s' % dte)

    return dte


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

    # Get todays date
    dte = get_date(1)

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


def load_fs_workbook():
    """ Load workbook """
    logger.info('Loading fs workbook')

    try:
        wb = pyxl.load_workbook(filename=SOURCE)

    except Exception as e:
        logger.exception('Workbook load error : %s' % e)
        quit()

    logger.info('Loading of fs complete')

    return wb


def get_forex_rates():
    """ Get forex rates """
    logger.info('\nGetting forex rates..')

    ws = wb[RATESHEET]

    row = 2

    while (ws.cell(row, 1).value):
        currency = ws.cell(row, 1).value
        # logger.info('Currency : %s' % currency)

        if currency[0: 3] == 'USD':
            logger.debug('USD URL : %s' % USDFXURL)
            page = requests.get(USDFXURL)
        elif currency[0: 3] == 'GBP':
            logger.debug('GBP URL : %s' % GBPFXURL)
            page = requests.get(GBPFXURL)
        elif currency[0: 3] == 'CAD':
            logger.debug('CAD URL : %s' % CADFXURL)
            page = requests.get(CADFXURL)
        elif currency[0: 3] == 'SGD':
            logger.debug('SGD URL : %s' % SGDFXURL)
            page = requests.get(SGDFXURL)

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


def get_security_price_from_yahoo(type, country, sheetname, tickercol, pricecol):
    """
    Get security price from Yahoo Finance
    Parameters:
        type - string; security type (EQ - Equity, UT - Unit Trust, CO - Commodities)
        country - string; listed security country (US - USA, SG - Singapore, IN - India)
        sheetname - string; input Excel sheet name
        tickercol - numeric; column in Excel sheet containing ticker
        pricecol - numeric; Excel column where ticker price is to be output
    """
    ws = wb[sheetname]

    row = 2

    # Assign url based on country
    if country == 'IN':
        srcurl = 'https://in.finance.yahoo.com/quote/'
    elif country == 'SG':
        srcurl = 'https://sg.finance.yahoo.com/quote/'
    elif country == 'US':
        srcurl = 'https://finance.yahoo.com/quote/'
    else:
        srcurl = 'https://finance.yahoo.com/quote/'

    logger.info('\nGetting stock prices from ' + srcurl)

    # Loop through the ticker sheet column, extract the price and assign to price column
    while (ws.cell(row, tickercol).value):
        ticker = ws.cell(row, tickercol).value

        if type == 'UT':
            if country == 'IN':
                yahoourl = srcurl + str(ticker) + '.BO?p=' + str(ticker) + '.BO&.tsrc=fin-srch-v1'
            elif country == 'SG':
                yahoourl = srcurl
        elif type == 'CO':
            yahoourl = srcurl + 'GC=F'
        else:
            if country == 'IN':
                if ticker == 'SRGHFL' or ticker == ' 500193':
                    yahoourl = srcurl + ticker + '.BO?p=' + ticker + '.BO&.tsrc=fin-srch'
                else:
                    yahoourl = srcurl + ticker + '.NS?p=' + ticker + '.NS&.tsrc=fin-srch'

            else:
                if country == 'SG' or country == 'US':
                    # yahoourl = srcurl + ticker + '?p=' + ticker + '&.tsrc=fin-srch'
                    yahoourl = srcurl + ticker + '?p=' + ticker

        logger.debug('URL : %s' % yahoourl)

        page = requests.get(yahoourl)

        soup = BeautifulSoup(page.content, 'lxml')

        # Price
        try:
            divtag = soup.find('div', class_='My(6px) Pos(r) smartphone_Mt(6px)')

        except Exception as e:
            logger.exception('Find error : %s' % e)
            row += 1
            continue

        try:
            price = divtag.div.span.text

        except Exception as e:
            logger.exception('Tag not found : %s' % e)
            row += 1
            continue

        ws.cell(row, pricecol).value = float(price.replace(',', ''))
        logger.info('%s / %s' % (ticker, ws.cell(row, pricecol).value))

        if (country == 'US') and (ticker not in EXCLUDE):
            # Fair Value
            try:
                valuediv = soup.findAll('div', class_='Fw(b) Fl(end)--m Fz(s) C($primaryColor')
                ws.cell(row, 24).value = valuediv[0].text

            except Exception as e:
                logger.exception('Find error : %s' % e)
                row += 1
                continue

        if (country == 'US') and (ticker not in EXCLUDE):
            # Pattern
            try:
                patdiv = soup.find('div', class_='W(1/4)--mobp W(1/2) IbBox')

            except Exception as e:
                logger.exception('Find error : %s' % e)
                row += 1
                continue

            try:
                ws.cell(row, 25).value = patdiv.div.span.text

            except Exception as e:
                logger.exception('Index error : %s' % e)
                row += 1
                continue

        if country == 'US':
            # Stats
            if ticker not in EXCLUDE:
                yahoourl = srcurl + ticker + '/key-statistics?p=' + ticker

                logger.debug('URL : %s' % yahoourl)

                page = requests.get(yahoourl)

                soup = BeautifulSoup(page.content, 'lxml')

                try:
                    tds = soup.findAll('td', class_='Ta(c) Pstart(10px) Miw(60px) Miw(80px)--pnclg Bgc($lv1BgColor) fi-row:h_Bgc($hoverBgColor)')

                except Exception as e:
                    logger.exception('Find error : %s' % e)
                    row += 1
                    continue

                try:
                    ws.cell(row, 15).value = float(tds[0].text)
                    ws.cell(row, 16).value = float(tds[1].text)
                    ws.cell(row, 17).value = float(tds[2].text)
                    ws.cell(row, 18).value = float(tds[3].text)
                    ws.cell(row, 19).value = float(tds[4].text)
                    ws.cell(row, 20).value = float(tds[5].text)
                    ws.cell(row, 21).value = float(tds[6].text)
                    ws.cell(row, 22).value = float(tds[7].text)
                    ws.cell(row, 23).value = float(tds[8].text)

                except Exception as e:
                    logger.exception('Stats index error : %s' % e)
                    row += 1
                    continue

            # Sector/Industry
            if (ws.cell(row, 2).value is None) or (ws.cell(row, 3).value is None):
                yahoourl = srcurl + ticker + '/profile?p=' + ticker

                logger.debug('URL : %s' % yahoourl)

                page = requests.get(yahoourl)

                soup = BeautifulSoup(page.content, 'lxml')

                try:
                    tds = soup.findAll('span', class_='Fw(600)')

                except Exception as e:
                    logger.exception('Find error : %s' % e)
                    row += 1
                    continue

                try:
                    ws.cell(row, 2).value = tds[0].text
                    ws.cell(row, 3).value = tds[1].text

                except Exception as e:
                    logger.exception('Find error : %s' % e)
                    row += 1
                    continue

        elif country == 'IN' and type == 'EQ':
            # Stats
            yahoourl = srcurl + ticker + '.NS' + '/key-statistics?p=' + ticker + '.NS'

            logger.debug('URL : %s' % yahoourl)

            page = requests.get(yahoourl)

            soup = BeautifulSoup(page.content, 'lxml')

            try:
                tds = soup.findAll('td', class_='Fw(500) Ta(end) Pstart(10px) Miw(60px)')

            except Exception as e:
                logger.exception('Find error : %s' % e)
                row += 1
                continue

            ws.cell(row, 14).value = float(tds[0].text)
            ws.cell(row, 15).value = float(tds[1].text)
            ws.cell(row, 16).value = float(tds[2].text)
            ws.cell(row, 17).value = float(tds[3].text)
            ws.cell(row, 18).value = float(tds[4].text)
            ws.cell(row, 19).value = float(tds[5].text)
            ws.cell(row, 20).value = float(tds[6].text)
            ws.cell(row, 21).value = float(tds[7].text)
            ws.cell(row, 22).value = float(tds[8].text)

            # Sector/Industry
            if (ws.cell(row, 2).value is None) or (ws.cell(row, 3).value is None):
                yahoourl = srcurl + ticker + '.NS' + '/profile?p=' + ticker + '.NS'

                logger.debug('URL : %s' % yahoourl)

                page = requests.get(yahoourl)

                soup = BeautifulSoup(page.content, 'lxml')

                try:
                    tds = soup.findAll('span', class_='Fw(600)')

                except Exception as e:
                    logger.exception('Find error : %s' % e)
                    row += 1
                    continue

                try:
                    ws.cell(row, 2).value = tds[0].text
                    ws.cell(row, 3).value = tds[1].text

                except Exception as e:
                    logger.exception('Find error : %s' % e)
                    row += 1
                    continue

        row += 1
        time.sleep(DELAY)


def get_singapore_ut_prices():
    """ Get Singapore UT price """
    logger.info('\nGetting SG UT price..')

    ws = wb[SGUTSHEET]

    try:
        page = requests.get(SGUTURL)

    except Exception as e:
        logger.exception('URL open error : %s' % e)
        quit()

    logger.debug('HTML : %s' % page)

    soup = BeautifulSoup(page.content, 'lxml')
    logger.debug('Soup : %s' % soup.body)

    panel = soup.find('a', string='Infinity US 500 Stock Index Fund SGD')
    logger.debug('Tag : %s' % panel)

    price = panel.next_element.next_element.text

    ws['L2'].value = float(price)
    logger.info('Infinity US 500 Stock Index Fund SGD Price : %s' % ws['L2'].value)


def get_crypto_prices():
    """ Get Crypto prices """
    logger.info('\nGetting Crypto prices..')

    ws = wb[CRYPTOSHEET]

    row = 2

    while (ws.cell(row, 2).value):
        cryptoticker = ws.cell(row, 2).value
        # logger.info('Crypto ticker : %s' % cryptoticker)

        cryptourl = 'https://min-api.cryptocompare.com/data/price?fsym=' + \
            cryptoticker + \
            '&tsyms=BTC,USD'
        logger.debug('Crypto URL : %s' % cryptourl)

        r = requests.get(cryptourl)
        parsed_json = r.json()

        try:
            ws.cell(row, 4).value = float(parsed_json['USD'])

        except Exception as e:
            logger.exception('Index error : %s' % e)
            continue

        logger.info('Ticker / Price : %s / %s' % (cryptoticker, ws.cell(row, 4).value))

        logger.debug('Row : %s' % row)

        row += 1
        time.sleep(DELAY)


def save_fs_workbook():
    """ save workbook """

    ws = wb[FSSHEET]

    # Get todays date
    dte = get_date(2)

    ws.cell(30, 3).value = dte

    logger.info('Saving FS workbook..')

    try:
        wb.save(SOURCE)

    except Exception as e:
        logger.exception('Workbook save error : %s' % e)
        quit()

    logger.info('Saved')


if __name__ == '__main__':
    logger = setup_logger()

    check_file_exists(SOURCE)

    backup_input_file()

    wb = load_fs_workbook()

    get_forex_rates()

    get_security_price_from_yahoo('EQ', 'US', USSHEET, 2, 9)
    get_security_price_from_yahoo('EQ', 'US', USWLSHEET, 1, 4)

    get_security_price_from_yahoo('EQ', 'IN', INSHEET, 2, 7)
    get_security_price_from_yahoo('EQ', 'IN', INWLSHEET, 1, 4)

    get_security_price_from_yahoo('UT', 'IN', INUTSHEET, 2, 12)

    get_security_price_from_yahoo('EQ', 'SG', SGSHEET, 2, 8)
    get_singapore_ut_prices()

    get_crypto_prices()

    save_fs_workbook()
    
