#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pytest
from stockapi import stocks

__author__ = "vinodvoa"
__copyright__ = "vinodvoa"
__license__ = "mit"

YAHOOURL1 = 'https://in.finance.yahoo.com/quote/' + 'SRGHFL' + \
    '.BO?p=' + 'SRGHFL' + '.BO&.tsrc=fin-srch-v1'

YAHOOURL2 = 'https://in.finance.yahoo.com/quote/' + 'PIDILITIND' + \
    '.NS?p=' + 'PIDILITIND' + '.NS&.tsrc=fin-srch-v1'

YAHOOURL3 = 'https://finance.yahoo.com/quote/' + 'D05.SI' + \
    '?p=' + 'D05.SI' + '&.tsrc=fin-srch'

YAHOOURL4 = 'https://finance.yahoo.com/quote/' + 'TCEHY' + \
    '?p=' + 'TCEHY' + '&.tsrc=fin-srch'

IEXURL = 'https://api.iextrading.com/1.0/stock/' + 'FB' + '/book'

CRYPTOURL = 'https://min-api.cryptocompare.com/data/price?fsym=' + \
    'BTC' + '&tsyms=BTC,USD'


@pytest.mark.parametrize('url, rc',
                         [
                             (stocks.GOLDURL, 200),
                             (stocks.SGUTURL, 200),
                             (stocks.USDFXURL, 200),
                             (stocks.CADFXURL, 200),
                             (stocks.GBPFXURL, 200),
                             (stocks.SGDFXURL, 200),
                             (stocks.SGYURL, 200),
                             (YAHOOURL1, 200),
                             (YAHOOURL2, 200),
                             (YAHOOURL3, 200),
                             (YAHOOURL4, 200),
                             (IEXURL, 200),
                             (CRYPTOURL, 200)
                         ]
                         )
def test_urls(url, rc):
    assert stocks.check_urls(url) == rc
