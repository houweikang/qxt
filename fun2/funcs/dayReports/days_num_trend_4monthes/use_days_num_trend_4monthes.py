#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/27 16:13:18
# @Author  : HouWk
# @Site    : 
# @File    : use_days_num_trend_4monthes.py
# @Software: PyCharm
from days_num_trend_4monthes.data_days_num_trend_4monthes import days_num_trend_4monthes
from fun_date import get_str_lastNmonth_firstday
from fun_report_chart import MakeReportOrChart
from use_win32com_exl_chart import chart_4monthes_trend


def reports_days_num_trend_4monthes(dq, app, start_date,final_date):
    dfs_list = days_num_trend_4monthes(start_date, final_date, dq)
    color = 'R'

    dfs_colege = dfs_list[0]
    for df in dfs_colege:
        dep = df.iloc[0, 0]
        sheet_name = '%s日创量趋势' % dep
        report_chart = MakeReportOrChart(app=app, df=df, sheet_name=sheet_name, color=color)
        # report
        report_chart.report(start_date, final_date)
        # chart
        chart_title = '%s%s-近%d个月每日创量总业绩趋势' % ('QX', dep, 4)
        report_chart.chart(chart_title=chart_title, chart_rng_list=[[3, 2]], chart_style=chart_4monthes_trend)

    dfs_region = dfs_list[1]
    for df in dfs_region:
        dep = ''.join(list(df.iloc[0, :2]))
        sheet_name = '%s日创量趋势' % dep
        report_chart = MakeReportOrChart(app=app, df=df, sheet_name=sheet_name, color=color)
        # report
        report_chart.report(start_date, final_date)
        # chart
        chart_dq = df.iloc[0, 1]
        chart_title = '%s%s-近%d个月每日创量总业绩趋势' % ('QX', chart_dq, 4)
        report_chart.chart(chart_title=chart_title, chart_rng_list=[[3, 3]], chart_style=chart_4monthes_trend)

    dfs_team = dfs_list[2]
    for df in dfs_team:
        dep = ''.join(list(df.iloc[0, :3]))
        sheet_name = '%s日创量趋势' % dep
        report_chart = MakeReportOrChart(app=app, df=df, sheet_name=sheet_name, color=color)
        # report
        report_chart.report(start_date, final_date)
        # chart
        chart_title = '%s%s-近%d个月每日创量总业绩趋势' % ('QX', dep, 4)
        report_chart.chart(chart_title=chart_title, chart_rng_list=[[3, 4]], chart_style=chart_4monthes_trend)