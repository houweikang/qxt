#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/12 9:42:08
# @Author  : HouWk
# @Site    : 
# @File    : os_walk.py
# @Software: PyCharm
import os
import pandas as pd


def walk(path):
    p_datas = []
    if not os.path.exists(path):
        return -1
    for root, dirs, names in os.walk(path):
        for filename in names:
            filepath = os.path.join(root, filename)  # 路径和文件名连接构成完整路径
            dt = '{}-{}-{}'.format(filename[:4], filename[4:6], filename[6:8])
            p_data = pd.read_excel(filepath)
            p_data['日期'] = dt
            p_data['日期'] = pd.to_datetime(p_data['日期'], format='%Y-%m-%d')
            p_datas.append(p_data)
    result_all = pd.concat(p_datas, sort=False)
    # result_all.to_excel('zhucejinqun202001.xlsx')


if __name__ == '__main__':
    no_work_file_path = r'e:\报表\所有数据信息\员工信息\保定员工信息\202004\20200420保定员工信息.xls'
    path = r'c:\Users\Administrator\Desktop\202005'
    walk(path)