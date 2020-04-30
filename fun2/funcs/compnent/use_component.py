#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/20 15:24:33
# @Author  : HouWk
# @Site    : 
# @File    : data_component.py
# @Software: PyCharm
from compnent.data_component import component_data
from fun_date import get_str_date, get_str_lastNmonth_firstday
from config import root_path, component_dqs
from fun_report_chart import MakeReportOrChart
from use_win32com_exl_chart import chart_component
from use_os import create_folder_date
from fun_win32com_exl import Excel


def component_report( final_date):
    start_date = get_str_lastNmonth_firstday(final_date, -3)  # 获取起始日期
    app = Excel()
    app.screen_updating(False)  # 关闭屏幕刷新
    app.workbooks_add()  # 创建excel wb
    for dq in component_dqs:
        df = component_data(start_date, final_date, dq)
        sheet_name = '%s近4月日人均分量' % dq
        report_chart = MakeReportOrChart(app=app, df=df, sheet_name=sheet_name, color='R')
        # report
        report_chart.report(start_date, final_date)
        # chart
        chart_title = '%s%s-近%d个月每日人均分量趋势' % ('QX', dq, 4)
        report_chart.chart(chart_title=chart_title, chart_style=chart_component)
    app.sheets_delete(1)
    app.sheets_select(1)
    # 保存文件
    path = create_folder_date(root_path, final_date)  # 创建目标文件夹
    str_date = get_str_date(final_date, '%Y%m%d')
    wb_name = '%s日报_人均分量' % str_date
    app.workbooks_save(wb_name, path)
    app.screen_updating(True)  # 关闭屏幕刷新  # 开启屏幕刷新

