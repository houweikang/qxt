#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/24 10:54:01
# @Author  : HouWk
# @Site    : 
# @File    : fun_titles.py
# @Software: PyCharm
from fun_date import get_str_date


def subtitle(*args):
    n = len(args)
    if n == 1:
        final_date = args[0]
        final_date = get_str_date(final_date, '%Y/%m/%d')
        dates = final_date
        return dates
    elif n == 2:
        start_date = args[0]
        start_date = get_str_date(start_date, '%Y/%m/%d')
        final_date = args[1]
        final_date = get_str_date(final_date, '%Y/%m/%d')
        dates = '%s-%s' % (start_date, final_date)
        return dates
    else:
        print('参数过多！')
        exit()


def titles(title, *args):
    t = [[title, '']]
    subt = subtitle(*args)
    subts = ['统计时间：', subt]
    t.append(subts)
    return t


if __name__ == '__main__':
    print(titles('123', '2020/1/1', '2020/4/1'))
    print(titles('123', '2020/4/1'))
