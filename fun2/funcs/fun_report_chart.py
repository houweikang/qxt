#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 17:11:08
# @Author  : HouWk
# @Site    : 
# @File    : fun_report_chart.py
# @Software: PyCharm
from fun_titles import titles
from fun_win32com_exl_range_style import RangeStyle


class MakeReportOrChart:

    def __init__(self, app, df, sheet_name, color):
        self.result_value = df.values
        self.result_value = list(self.result_value)
        self.col_names = list(df.columns)
        self.sheet_name = sheet_name
        self.app = app
        self.sht = app.sheets_add(sheet_name)  # 添加sheet
        app.sheet_tab_color(sheet_name, color)

    def report(self, start_date, final_date):
        tles = titles(self.sheet_name, start_date, final_date)  # 标题和副标题
        # 数据写入excel
        self.app.write_data(self.sht, "a1", tles)
        self.app.write_data(self.sht, "a3", self.col_names)
        self.app.write_data(self.sht, 'a4', self.result_value)
        # 调整数据区域格式
        rngstyle = RangeStyle(self.sht)
        rngstyle.none_gridlines()  # 无网格线
        # 赋值
        rg = rngstyle.range
        cols_rg = rg.Rows(3)
        self.rs = rg.Rows.Count
        self.cs = rg.Columns.Count
        data_rg = rngstyle.sheet.Range(rg.Cells(4, 1), rg.Cells(self.rs, self.cs))
        # 设置格式
        rngstyle.alignment()  # 居中对齐
        rngstyle.fontname()  # 雅黑字体
        rngstyle.title_style()
        rngstyle.subtitle_style()
        rngstyle.cols_style(cols_rg, False)
        rngstyle.data_style(data_rg, False)

    def chart(self, chart_title, chart_style, chart_rng_list=None, chart_name=None):
        # chart_rng_list=[[],[]]
        # 调整chart
        try:
            if chart_rng_list:
                if len(chart_rng_list) == 2:
                    cel1 = self.sht.Cells(chart_rng_list[0][0], chart_rng_list[0][1])
                    cel2 = self.sht.Cells(chart_rng_list[1][0], chart_rng_list[1][1])
                    chart_rng = self.sht.Range(cel1, cel2)
                elif len(chart_rng_list) == 1:
                    cel1 = self.sht.Cells(chart_rng_list[0][0], chart_rng_list[0][1])
                    cel2 = self.sht.Cells(self.rs, self.cs)
                    chart_rng = self.sht.Range(cel1, cel2)
            else:
                chart_rng = self.sht.Range(self.sht.Cells(3, 1), self.sht.Cells(self.rs, self.cs))
            if not chart_name:
                chart_name = self.sheet_name
            chart_style(chart_rng, chart_title, chart_name)  # 调整数据显示格式
        except Exception as error:
            print('error:', error)
