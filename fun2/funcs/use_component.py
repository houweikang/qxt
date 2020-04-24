#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/20 15:24:33
# @Author  : HouWk
# @Site    : 
# @File    : fun_component.py
# @Software: PyCharm
from db_QXT import operate_db
import pandas as pd
from fun_date import get_str_date
import numpy as np
import matplotlib.pyplot as plt
from config import root_path
from use_os import create_folder_date
from fun_win32com_exl import workbooks_add,write_data
from fun_win32com_exl import activesheet  #实验用
from fun_titles import titles
from fun_win32com_exl_range_style import component_style


def component_data(start_date, final_date, dq):
    sql = '''SELECT cast([提交时间] as date) 日期,[课程顾问-员工号] 工号
        FROM [QXT].[dbo].[Tg] 
        where cast([提交时间] as date) between '{}' and '{}' and [课程顾问-所属地区] = '{}'
        '''.format(start_date, final_date, dq)
    data = operate_db(sql)

    # 获取每日顾问人数
    unique_consultant = data.drop_duplicates()
    consultant_num = unique_consultant.groupby('日期').count()
    consultant_num.rename(columns={'工号': '顾问人数'}, inplace=True)

    # 获取每日分量
    component_num = data.groupby('日期').count()
    component_num.rename(columns={'工号': '分量'}, inplace=True)

    # 获取日人均分量
    avg_cs_cp = pd.merge(consultant_num, component_num, on='日期')
    avg_cs_cp['日人均分量'] = avg_cs_cp.分量 / avg_cs_cp.顾问人数
    avg_cs_cp['日人均分量'] = avg_cs_cp['日人均分量'].map(lambda x: '%.0f' % x)

    # reindex日期索引列
    new_index = list(pd.date_range(start_date, final_date))
    avg_cs_cp = avg_cs_cp.reindex(new_index, fill_value=0)
    avg_cs_cp.reset_index(level=0, inplace=True)
    avg_cs_cp['日期'] = pd.to_datetime(avg_cs_cp.日期)
    avg_cs_cp['月'] = avg_cs_cp.日期.dt.month
    avg_cs_cp['日'] = avg_cs_cp.日期.dt.day

    result = avg_cs_cp[['月', '日', '日人均分量']]
    result.set_index(['月', '日'], inplace=True)
    result = result.unstack()  # 转置
    result.reset_index(inplace=True)
    result['月'] = result['月'].map(lambda x: '%d月' % x)  # 改变index
    result.columns = ['%s日' % col[1] for col in result.columns]  # 改变cols
    result = result.fillna('') #nan替换为空值

    col_names = list(result.columns)
    col_names[0] = '月'
    result.columns = col_names  # 行名

    path  = create_folder_date(root_path, final_date)  # 创建目标文件夹

    sheet_name = '%s近4月日人均分量' % dq
    # sht = workbooks_add(sheet_name) # 创建excel sheet1
    sht = activesheet() # 创建excel sheet1

    tles = titles(sheet_name,start_date,final_date) #标题和副标题

    # 数据写入excel
    write_data(sht,"a1",tles)
    write_data(sht,"a3",col_names)
    write_data(sht,'a4',list(result.values))

    component_style(sht)  #调整数据显示格式



    # writer = pd.ExcelWriter(r'c:\Users\Administrator\Desktop\%s' % wb_name)
    # result.to_excel(writer, sheet_name=sheet_name, index=False)
    # writer.save()
    # sht = xw.sheets.active
    # print(result.columns)
    # sht[0, 0].value = col_names
    # sht[1, 0].value = result.values


if __name__ == '__main__':
    ComD = component_data('2020/3/1', '2020/4/22', '燕郊')
    # print(ComD)

