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
from fun_str import replace_0


def get_date_date(date):
    #获取日期类型 日期
    if isinstance(date,str):
        date_date = parse(date)
    elif isinstance(date,date):
        date_date = date
    return date_date

def get_str_date(input_date=(datetime.date.today()+datetime.timedelta(-1)),format = '%Y/%m/%d'):
    '''
    :param date: 字符串类型日期
    :return: 字符串类型日期 '2020/5/6'
    '''
    if isinstance(input_date,str):
        input_date = get_date_date(input_date)
    str_date = input_date.strftime(format)
    str_date = replace_0(str_date)
    return str_date

def get_str_lastNmonth_firstday(date,n=0,format = '%Y/%m/%d'):
    '''
    :param date: 字符串类型日期
    :return: 当月第一天 字符串类型日期
    '''
    date_date = get_date_date(date)
    days = date_date.day
    date_currentmonth_firstday = date_date + datetime.timedelta(1-days)
    date_lastNmonth_firstday = date_currentmonth_firstday + relativedelta(months=n)
    str_currentmonth_firstday = get_str_date(date_lastNmonth_firstday,format)
    return str_currentmonth_firstday

def get_str_lastNmonth_endday(date,n=0,format = '%Y/%m/%d'):
    '''
    :param date: 字符串类型日期
    :return: 当月第一天 字符串类型日期
    '''
    date_date = get_date_date(date)
    days = date_date.day
    date_currentmonth_firstday = date_date + datetime.timedelta(1-days)
    date_lastNmonth_firstday = date_currentmonth_firstday + relativedelta(months=(n+1))
    date_lastNmonth_endday = date_lastNmonth_firstday+datetime.timedelta(-1)
    date_lastNmonth_endday = get_str_date(date_lastNmonth_endday,format)
    return date_lastNmonth_endday

def get_str_lastNmonth_sameday(date,n=0,format = '%Y/%m/%d'):
    '''
    :param date: 字符串类型日期
    :return: 当月第一天 字符串类型日期
    '''
    date_date = get_date_date(date)
    days = date_date.day
    date_currentmonth_firstday = date_date + datetime.timedelta(1-days)
    date_lastNmonth_firstday = date_currentmonth_firstday + relativedelta(months=n)
    date_lastNmonth_sameday = date_lastNmonth_firstday+datetime.timedelta(days-1)
    date_lastNmonth_sameday = get_str_date(date_lastNmonth_sameday,format)
    return date_lastNmonth_sameday

if __name__ == '__main__':
    print(get_date_date('2020/4/4'))
    print(get_str_lastNmonth_firstday('2020/4/4',-3))
    print(get_str_lastNmonth_endday('2020/5/4',0))
    print(get_str_lastNmonth_sameday('2020/5/4',-1))