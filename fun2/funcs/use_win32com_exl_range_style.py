#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 11:13:08
# @Author  : HouWk
# @Site    : 
# @File    : use_win32com_exl_range_style.py
# @Software: PyCharm
# from fun_win32com_exl_range_style import RangeStyle
#
#
# def component_style(sht=None):
#     # 无网格线
#     rngstyle = RangeStyle(sht)
#     rngstyle.none_gridlines()
#     # 赋值
#     rg = rngstyle.range
#     cols_rg = rg.Rows(3)
#     rc = rg.Rows.Count
#     cc = rg.Columns.Count
#     data_rg = rngstyle.sheet.Range(rg.Cells(4, 1), rg.Cells(rc, cc))
#     # 设置格式
#     rngstyle.alignment()  # 居中对齐
#     rngstyle.fontname()  # 雅黑字体
#     rngstyle.title_style()
#     rngstyle.subtitle_style()
#     rngstyle.cols_style(cols_rg, False)
#     rngstyle.data_style(data_rg, False)