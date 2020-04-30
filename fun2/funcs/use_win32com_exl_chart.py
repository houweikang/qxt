#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 10:37:21
# @Author  : HouWk
# @Site    : 
# @File    : use_win32com_exl_chart.py
# @Software: PyCharm
from config import component_RGB, RGBs
from fun_win32com_exl_chart import ExcelChart, font_style, postion
from win32com.client import constants as c  # 旨在直接使用VBA常数


def chart(data_rng, chart_name, chart_size, chart_type=c.xlLine,
          chart_plotby=c.xlRows, no_gridline=True, chart_style_num=None,
          chart_title=False, title_font_name='Microsoft YaHei UI', title_font_size=24,
          has_data_table=False, table_font_name='Microsoft YaHei UI', table_font_size=12,
          has_legend=False, legend_font_name='Microsoft YaHei UI',
          legend_font_size=12, legend_position=c.xlLegendPositionTop,
          y_stick_style=False, y_stick_font_name='Microsoft YaHei UI', y_stick_font_size=12,
          series_style=False, series_lines_smooth=False,
          series_colors=False, series_lines_color_list=None, series_reverse=True,
          ):
    # chart_size = (0, 170, 1200, 500)
    chart = ExcelChart(chart_name, chart_size, chart_type)
    chart.data(data_rng, ChartPlotBy=chart_plotby)

    if no_gridline:
        chart.no_gridline()  # 无网格线

    if chart_style_num:
        chart.chartstyle(chart_style_num)  # 设置style

    if chart_title:
        title = chart.title(chart_title)  # title
        title_font = title.Font
        font_style(title_font, title_font_name, title_font_size)

    if has_data_table:
        dt = chart.data_table()  # datatable
        dt_font = dt.Font
        font_style(dt_font, table_font_name, table_font_size)

    if has_legend:
        legend = chart.legend()  # legend
        legend_font = legend.Font
        font_style(legend_font, legend_font_name, legend_font_size)
        postion(legend, legend_position)

    if y_stick_style:
        y_stick = chart.ticklabel()  # Y轴
        y_stick_font = y_stick.Font
        font_style(y_stick_font, y_stick_font_name, y_stick_font_size)

    if series_style:
        series_count = chart.series_count()  # series
        if series_colors:
            position_list = list(range(1, series_count + 1))
            if series_reverse:
                series_lines_color_list = series_lines_color_list[::-1]
                position_list = position_list[::-1]

            for n,i in enumerate(position_list):
                chart.series_style(i, series_lines_smooth, series_lines_color_list[n])
    chart.select()
    return chart


def chart_component(data_rng, chart_title, chart_name):
    return chart(data_rng, chart_name, chart_size=(0, 170, 1200, 500),
                 chart_plotby=c.xlRows, no_gridline=True, chart_style_num=233,
                 chart_title=chart_title, title_font_name='Microsoft YaHei UI', title_font_size=24,
                 has_data_table=True, table_font_name='Microsoft YaHei UI', table_font_size=12,
                 has_legend=True, legend_font_name='Microsoft YaHei UI',
                 legend_font_size=12, legend_position=c.xlLegendPositionTop,
                 y_stick_style=True, y_stick_font_name='Microsoft YaHei UI', y_stick_font_size=12,
                 series_style=True, series_lines_smooth=True,
                 series_colors=True, series_lines_color_list=component_RGB, series_reverse=True,
                 )

def chart_4monthes_trend(data_rng, chart_title, chart_name):
     chart_4monthes=chart(data_rng, chart_name, chart_size=(0, 170, 1400, 500),
                 chart_plotby=c.xlRows, no_gridline=True, chart_style_num=233,
                 chart_title=chart_title, title_font_name='Microsoft YaHei UI', title_font_size=24,
                 has_data_table=True, table_font_name='Microsoft YaHei UI', table_font_size=12,
                 has_legend=True, legend_font_name='Microsoft YaHei UI',
                 legend_font_size=12, legend_position=c.xlLegendPositionTop,
                 y_stick_style=True, y_stick_font_name='Microsoft YaHei UI', y_stick_font_size=12,
                 series_style=True, series_lines_smooth=True,
                 series_colors=True, series_lines_color_list=RGBs, series_reverse=True,
                 )
     chart_4monthes
