#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 17:00:29
# @Author  : HouWk
# @Site    :
# @File    : use_day_reports.py
# @Software: PyCharm
from days_num_trend_4monthes.use_days_num_trend_4monthes import reports_days_num_trend_4monthes
from enrollment_and_groupEntry.use_enrollment_and_GroupEntry_rate import er_and_GE_R_report
from fun_date import get_str_date, get_str_lastNmonth_firstday
from config import root_path, day_reports_list
from use_os import create_folder_date
from fun_win32com_exl import Excel


def day_reports(final_date):
    app = Excel()
    app.screen_updating(False)  # 关闭屏幕刷新
    for dq in day_reports_list:
        app.workbooks_add()  # 创建excel wb
        start_date = get_str_lastNmonth_firstday(final_date, 0)  # 获取起始日期
        start_date_3 = get_str_lastNmonth_firstday(final_date, -3)  # 获取起始日期

        # 报表
        # er_and_GE_R_report(dq, app=app, start_date=start_date, final_date=final_date, T_or_G='T')  # 学院、地区、运营部 注册与进群

        reports_days_num_trend_4monthes(dq, app, start_date_3, final_date)

        # er_and_GE_R_report(dq, app=app, start_date=start_date, final_date=final_date, T_or_G='G')  # 小组 注册与进群

        app.sheets_delete(1)
        app.sheets_select(1)
        # 保存文件
        path = create_folder_date(root_path, final_date)  # 创建目标文件夹
        str_date = get_str_date(final_date, '%Y%m%d')
        wb_name = '%s%s日报' % (str_date, dq)
        app.workbooks_save(wb_name, path)
    app.screen_updating(True)  # 开启屏幕刷新


if __name__ == '__main__':
    day_reports('2020/4/29')
