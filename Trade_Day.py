#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
Created on Wed Jul 10 11:04:25 2019

@author: arda
"""

import datetime

f = open('trade_date.txt', 'w')

today = datetime.date.today()
oneday = datetime.timedelta(days=1)
start_day = today - oneday*9
stop_day = today + oneday*174

day = start_day


while day <= stop_day:
    if day.weekday() != 5 and day.weekday() != 6: #如果非假日
        f.write(str(day) + '\n')
    else:
        pass
    
    day += oneday