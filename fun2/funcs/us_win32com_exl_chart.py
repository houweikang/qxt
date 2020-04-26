#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 10:37:21
# @Author  : HouWk
# @Site    : 
# @File    : us_win32com_exl_chart.py
# @Software: PyCharm
from fun_win32com_exl_chart import ExcelChart, font_style, postion
from win32com.client import constants as c  # 旨在直接使用VBA常数

RGBs = [(255, 255, 0), (0, 112, 192), (0, 176, 80), (255, 0, 0), (255, 255, 255)]


# 黄 蓝 绿 红 白

def chart_component(data_rng,chart_title,chart_name):
    component_RGB = RGBs[:4]

    chart = ExcelChart(chart_name, (0, 170, 1200, 500), c.xlLine)
    chart.data(data_rng, ChartPlotBy=c.xlRows)

    chart.no_gridline()  # 无网格线

    chart.chartstyle(233)  # 设置style

    title = chart.title(chart_title)  # title
    title_font = title.Font
    font_style(title_font, 'Microsoft YaHei UI', 24)

    dt = chart.data_table()  # datatable
    dt_font = dt.Font
    font_style(dt_font, 'Microsoft YaHei UI', 12)

    legend = chart.legend()  # legend
    legend_font = legend.Font
    font_style(legend_font, 'Microsoft YaHei UI', 12)
    postion(legend, c.xlLegendPositionTop)

    y_stick = chart.ticklabel()  # Y轴
    y_stick_font = y_stick.Font
    font_style(y_stick_font, 'Microsoft YaHei UI', 12)

    series_count = chart.series_count()  # series
    n = 0
    for i in range(series_count, 0, -1):  # series 倒序 ，component_RGB倒序输入
        chart.series_style(i, True, component_RGB[::-1][n])
        n += 1

    chart.select()


# if __name__ == '__main__':
#     chart_component('a')