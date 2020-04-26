#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/20 15:24:33
# @Author  : HouWk
# @Site    : 
# @File    : fun_component.py
# @Software: PyCharm
from fun_component import component_data
from fun_date import get_str_date
import numpy as np
from config import root_path, component_dqs
from us_win32com_exl_chart import chart_component
from use_os import create_folder_date
from fun_win32com_exl import Excel
from fun_titles import titles
from use_win32com_exl_range_style import component_style


def component_report(start_date, final_date):
    app = Excel()
    app.screen_updating(False)  # 关闭屏幕刷新
    wb = app.workbooks_add()  # 创建excel wb
    for dq in component_dqs:
        result_value, col_names = component_data(start_date, final_date, dq)
        sheet_name = '%s近4月日人均分量' % dq
        sht = app.sheets_add(sheet_name)  # 添加sheet
        app.sheet_tab_color(sheet_name, 'R')
        tles = titles(sheet_name, start_date, final_date)  # 标题和副标题
        # 数据写入excel
        app.write_data(sht, "a1", tles)
        app.write_data(sht, "a3", col_names)
        app.write_data(sht, 'a4', result_value)
        # 调整数据区域格式
        component_style(sht)
        # 调整chart
        result_value = np.array(result_value)
        shape = result_value.shape
        if len(shape) == 1:
            rs = 1
            cs = shape[0]
        elif len(shape) == 2:
            rs = shape[0]
            cs = shape[1]
        try:
            chart_rng = sht.Range(sht.Cells(3, 1), sht.Cells(3, 1).GetOffset(rs, cs - 1))
            chart_title = '%s%s-近%d个月每日人均分量趋势' % ('QX', dq, 4)
            chart_name = sheet_name
            chart_component(chart_rng, chart_title, chart_name)  # 调整数据显示格式
        except Exception as error:
            print('error:', error)

    app.sheets_delete(1)
    app.sheets_select(1)
    # 保存文件
    path = create_folder_date(root_path, final_date)  # 创建目标文件夹
    str_date = get_str_date(final_date, '%Y%m%d')
    wb_name = '%s日报_人均分量' % str_date
    app.workbooks_save(wb_name, path)
    app.screen_updating(True)  # 关闭屏幕刷新  # 开启屏幕刷新
