#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 16:46:27
# @Author  : HouWk
# @Site    : 
# @File    : use_enrollment_and_GroupEntry_rate.py
# @Software: PyCharm
from enrollment_and_groupEntry.data_erollment_and_GroupEntry_rate import er_and_ge_rate_data
from fun_report_chart import MakeReportOrChart


def er_and_GE_R_report(dq, app, start_date, final_date, T_or_G='T'):
    # group_data, team_data, region_data, colege_data = er_and_ge_rate_data(dq)
    df_list = list(er_and_ge_rate_data(dq))
    if T_or_G == 'T':
        df_list = df_list[1:][::-1]
        color = 'R'
    elif T_or_G == 'G':
        df_list = df_list[0:1]
        color = 'G'
    for df in df_list:
        dep = df.columns[-6]
        sheet_name = '%s推广进群率与注册率统计' % dep
        report_chart = MakeReportOrChart(app=app, df=df, sheet_name=sheet_name, color=color)
        # report
        report_chart.report(start_date, final_date)
