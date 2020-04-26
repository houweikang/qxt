#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/16 9:53:58
# @Author  : HouWk
# @Site    : 
# @File    : fun_date.py
# @Software: PyCharm
from dateutil.parser import parse
import datetime
from dateutil.relativedelta import relativedelta


def get_date_date(date):
    #获取日期类型 日期
    date_date = parse(date)
    return date_date

def get_str_date(input_date=(datetime.date.today()+datetime.timedelta(-1)),format = '%Y/%m/%d'):
    '''
    :param date: 字符串类型日期
    :return: 字符串类型日期 '2020/05/06'
    '''
    if isinstance(input_date,str):
        input_date = get_date_date(input_date)
    str_date = input_date.strftime(format)
    return str_date

def get_str_lastNmonth_firstday(date,n=0):
    '''
    :param date: 字符串类型日期
    :return: 当月第一天 字符串类型日期
    '''
    date_date = get_date_date(date)
    days = date_date.day
    date_currentmonth_firstday = date_date + datetime.timedelta(1-days)
    date_lastNmonth_firstday = date_currentmonth_firstday + relativedelta(months=n)
    str_currentmonth_firstday = date_lastNmonth_firstday.strftime('%Y/%m/%d')
    return str_currentmonth_firstday


if __name__ == '__main__':
    print(get_date_date('2020/4/4'))
    print(get_str_lastNmonth_firstday('2020/4/4',-3))
