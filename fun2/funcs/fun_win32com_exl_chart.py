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
        font = self.chart_chart.ChartTitle.Font
        font.Name = ''

    def lable(self, position=None):
        '''
        :param position: msoElementLegendTop ,msoElementLegendCenter
                        msoElementLegendBottom,msoElementLegendLeft,msoElementLegendRight

                        msoElementLegendNone
        :return:
        '''
        if not position:
            position=c.msoElementDataLabelTop
        self.chart_chart.ApplyDataLabels(ShowValue=True)
        self.chart_chart.SetElement(position)

    # def data_table(self, arg=c.msoElementDataTableWithLegendKeys):
    #     '''
    #     :param arg: msoElementDataTableNone
    #     :return:
    #     '''
    #     self.chart_chart.SetElement(arg)

    def other_type(self, num, type=c.xlLine):
        self.chart_chart.FullSeriesCollection(num).ChartType = type

    def AxisGroup(self, num):
        self.chart_chart.FullSeriesCollection(num).AxisGroup = 2

    def font(self, font_name='Microsoft YaHei UI'):
        self.chart_chart.ChartTitle.Font.Name = font_name
        self.chart_chart.Legend.Font.Name = font_name
        self.chart_chart.Datatable.Font.Name = font_name

    def no_gridline(self):
        self.chart_chart.SetElement(c.msoElementPrimaryValueGridLinesNone)

    def select(self):
        self.chart.Select()


def main():
    chart = ExcelChart('a', (0, 170, 1200, 500), c.xlLine)
    # chart.type(c.xlLine)
    chart.data(rng=chart.sheet.Range("a3:Af5"), ChartPlotBy=c.xlRows)
    # chart.no_gridline()
    # chart.font()
    # chart.data_table()
    chart.title('haha')
    # chart.lable()
    chart.select()


if __name__ == "__main__":
    main()
