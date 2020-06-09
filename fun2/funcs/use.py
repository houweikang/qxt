#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/6/8 14:09:36
# @Author  : HouWk
# @Site    : 
# @File    : use.py
# @Software: PyCharm
import datetime
from cls_reports import Reports
from fun_box import inputbox


def my_morning_evening_report():
    now_hour = datetime.datetime.now().hour
    if now_hour <= 17:
        report_datetime = (datetime.date.today() - datetime.timedelta(days=1)).strftime('%Y-%m-%d') + ' 0-24'
    else:
        report_datetime = (datetime.date.today()).strftime('%Y-%m-%d') + ' 0-17'
    report_datetime1 = inputbox(mylabel='晨晚报-请输入需统计日期和时间段', default=report_datetime)
    report_date, hours = report_datetime1.split()
    start_hour, end_hour = hours.split('-')
    Reports(final_date=report_date).reports_morning_evening(start_hour=int(start_hour), end_hour=int(end_hour))


def my_day_report():
    report_date = (datetime.date.today() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    report_date_str = inputbox(mylabel='日报-请输入需统计日期', default=report_date)
    cls_reports = Reports(final_date=report_date_str)
    cls_reports.report_component()  # 人均分量
    cls_reports.reports_day()  # 日报


if __name__ == '__main__':
    my_day_report()
    my_morning_evening_report()
