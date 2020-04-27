#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 14:20:45
# @Author  : HouWk
# @Site    : 
# @File    : report_day.py
# @Software: PyCharm
from fun_box import inputbox, only_ok
from fun_date import get_str_date, get_str_lastNmonth_firstday
from compnent.use_component import component_report


def report_day():
    final_date = get_str_date()
    final_date = inputbox('请输入统计时间：', final_date)  # 获取终止日期
    start_date = get_str_lastNmonth_firstday(final_date, -3)  # 获取起始日期

    component_report(start_date, final_date)

    only_ok()


if __name__ == '__main__':
    report_day()