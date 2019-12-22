#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pytest
from stockapi import stocks

__author__ = "vinodvoa"
__copyright__ = "vinodvoa"
__license__ = "mit"

def test_check_file_exsits():
    assert stocks.check_file_exists(stocks.filename) == True
