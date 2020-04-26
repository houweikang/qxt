#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/24 9:33:13
# @Author  : HouWk
# @Site    : 
# @File    : fun_win32com_exl_chart.py
# @Software: PyCharm
import win32com.client
from win32com.client import constants as c  # 旨在直接使用VBA常数

excel = win32com.client.gencache.EnsureDispatch("Excel.Application")


class ExcelChart:
    def __init__(self, chart_name, position_and_size, chart_type=c.xlColumnClustered, sht=None):
        '''
        c.xlColumnClustered -- 柱形图
        c.xlLineMarkers -- 折线散点图
        c.xlLine -- 折线图
        '''
        if sht:
            self.sheet = sht
        else:
            self.sheet = excel.ActiveSheet

        try:
            self.sheet.ChartObjects(chart_name).Delete()
        except:
            pass
        finally:
            self.chart = self.sheet.Shapes.AddChart(chart_type,
                                                    position_and_size[0],
                                                    position_and_size[1],
                                                    position_and_size[2],
                                                    position_and_size[3])
            self.chart.Name = chart_name
        self.chart_chart = self.chart.Chart

    def data(self, rng=None, ChartPlotBy=c.xlColumns):
        '''
        :param rng:
        :param ChartPlotBy: c.xlRows  or  c.xlColumns
        :return:
        '''
        if rng:
            rng = rng
        else:
            rng = self.sheet.UsedRange

        self.chart_chart.SetSourceData(Source=rng, PlotBy=ChartPlotBy)

    def title(self, name):
        self.chart_chart.HasTitle = True
        self.chart_chart.ChartTitle.Text = name
        return self.chart_chart.ChartTitle

    def lable(self):
        self.chart_chart.HasDataLabel = True
        self.chart_chart.DataLabel.Text = "Saturday"

    def legend(self):
        self.chart_chart.HasLegend = True
        return self.chart_chart.Legend

    def data_table(self):
        self.chart_chart.HasDataTable = True
        return self.chart_chart.DataTable

    def other_type(self, num, type=c.xlLine):
        self.chart_chart.FullSeriesCollection(num).ChartType = type

    def AxisGroup(self, num):
        self.chart_chart.FullSeriesCollection(num).AxisGroup = 2

    def font(self, font_name='Microsoft YaHei UI'):
        self.chart_chart.ChartTitle.Font.Name = font_name
        self.chart_chart.Legend.Font.Name = font_name
        self.chart_chart.Datatable.Font.Name = font_name

    def ticklabel(self):
        return self.chart_chart.Axes(c.xlValue).TickLabels

    def chartstyle(self, style_num=233):
        # Y轴文字
        # self.chart.ChartStyle = style_num
        self.chart_chart.ChartStyle = style_num

    def select(self):
        self.chart.Select()

    def no_gridline(self):
        self.chart_chart.Axes(c.xlValue).HasMajorGridlines = False

    def series_count(self):
        # return self.chart_chart.FullSeriesCollection.Count
        return len(self.chart_chart.FullSeriesCollection())

    def series_style(self, s_num, smooth=False, line_color=None):
        self.series = self.chart_chart.FullSeriesCollection(s_num)
        self.series.Smooth = smooth

        # RGB(255, 255, 0) 黄
        # RGB(0, 112, 192) 蓝
        # RGB(0, 176, 80) 绿
        # RGB(255, 0, 0) 红
        # RGB(255, 255, 255) 白

        if line_color:  # RGB
            self.series.Format.Line.ForeColor.RGB = rgbToInt(line_color)


def rgbToInt(rgb):
    colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
    return colorInt


def postion(obj, arg):
    obj.Position = arg


def font_style(font, name=None, size=None, bold=None):
    if name:
        font.Name = name
    if size:
        font.Size = size
    if bold:
        font.Bold = bold

# if __name__ == "__main__":
