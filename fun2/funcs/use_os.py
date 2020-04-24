#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/24 9:48:34
# @Author  : HouWk
# @Site    : 
# @File    : use_os.py
# @Software: PyCharm
from fun_os import create_folder
from datetime import datetime
import dateutil.parser


# 创建目标文件夹
# 包含：分量报表
def create_folder_date(path, date):
    try:
        date = dateutil.parser.parse(date)
        date_fmt = datetime.strftime(date, '%Y%m%d')
        path_year = date_fmt[:4]
        path_month = date_fmt[:6]
        path_day = date_fmt
        path = '''%s/%s/%s/%s/''' % (path, path_year, path_month, path_day)
        create_folder(path)
        return path
    except ValueError:
        print('未创建路径！')