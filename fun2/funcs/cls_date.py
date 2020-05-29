#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/12 10:28:06
# @Author  : HouWk
# @Site    : 
# @File    : cls_date.py
# @Software: PyCharm
from dateutil.parser import parse
import datetime
from dateutil.relativedelta import relativedelta
from fun_str import replace_0


class MyDate:
    def __init__(self, inputdate=datetime.date.today() + datetime.timedelta(-1)):
        if isinstance(inputdate, str):
            self.date_date = parse(inputdate)
        elif isinstance(inputdate, datetime.date):
            self.date_date = inputdate

    def get_date_current_month_firstday(self):
        '''
        :return: 当月第一天 日期类型
        '''
        days = self.date_date.day
        self.date_currentmonth_firstday = self.date_date + datetime.timedelta(1 - days)
        return self.date_currentmonth_firstday

    def get_date_Nmonthes_firstday(self, n=0):
        date_currentmonth_firstday = self.get_date_current_month_firstday()
        date_next_Nmonthes_firstday = date_currentmonth_firstday + relativedelta(months=n)
        return date_next_Nmonthes_firstday

    def get_date_Nmonthes_endday(self, n=0):
        date_next_Nmonthes_firstday = self.get_date_Nmonthes_firstday(n=n + 1)
        date_Nmonthes_endday = date_next_Nmonthes_firstday + datetime.timedelta(-1)
        return date_Nmonthes_endday

    def get_date_Nmonth_sameday(self, n=0,samemonth=True):
        '''
        :param date: 字符串类型日期
        :return: 当月第一天 字符串类型日期
        '''
        days = self.date_date.day
        date_Nmonthes_firstday = self.get_date_Nmonthes_firstday(n=n)
        date_Nmonthes_sameday = date_Nmonthes_firstday + datetime.timedelta(days - 1)
        if samemonth:
            while  date_Nmonthes_sameday.month != date_Nmonthes_firstday.month:
                date_Nmonthes_sameday = date_Nmonthes_sameday + datetime.timedelta(- 1)
        return date_Nmonthes_sameday


def get_str_date(date_date, format='%Y/%m/%d'):
    if isinstance(date_date, str):
        date_date = parse(date_date)
    elif isinstance(date_date, datetime.date):
        date_date = date_date
    str_date = date_date.strftime(format)
    str_date = replace_0(str_date)
    return str_date


if __name__ == '__main__':
    dt = MyDate('2020/5/31')
    print(dt.get_date_current_month_firstday())
    print(dt.get_date_Nmonth_sameday(n=-3))
    print(dt.get_date_Nmonthes_firstday(n=2))
    print(dt.get_date_Nmonthes_endday(n=-1))
