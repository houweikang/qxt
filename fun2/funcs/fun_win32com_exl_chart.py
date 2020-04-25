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



    def select(self):
        self.chart.Select()

    def no_gridline(self):
        self.chart_chart.Axes(c.xlValue).HasMajorGridlines = False


def postion(obj, arg):
    obj.Position = arg


def font_style(font, name=None, size=None, bold=None):
    if name:
        font.Name = name
    if size:
        font.Size = size
    if bold:
        font.Bold = bold


def main():
    chart = ExcelChart('a', (0, 170, 1200, 500), c.xlLine)
    chart.data(rng=chart.sheet.Range("a3:af7"), ChartPlotBy=c.xlRows)

    # chart.no_gridline()

    dt = chart.data_table()
    dt_font = dt.Font
    font_style(dt_font, 'Microsoft YaHei UI', 12, True)

    title = chart.title('haha')
    title_font = title.Font
    font_style(title_font, 'Microsoft YaHei UI', 20, True)

    legend = chart.legend()
    legend_font = legend.Font
    font_style(legend_font, 'Microsoft YaHei UI', 20, True)
    postion(legend, c.xlLegendPositionTop)

    chart.select()



if __name__ == "__main__":
    main()
