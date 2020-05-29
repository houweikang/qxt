#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/24 9:33:13
# @Author  : HouWk
# @Site    : 
# @File    : fun_win32com_exl_chart.py
# @Software: PyCharm
import win32com.client
from win32com.client import constants as c  # 旨在直接使用VBA常数
from config import component_RGB, RGBs


class ExcelChart:
    def __init__(self,
                 chart_name,
                 chart_size,
                 chart_type,
                 sht=None
                 ):
        '''
        c.xlColumnStacked -- 柱形堆积图
        c.xlColumnClustered -- 柱形图
        c.xlLineMarkers -- 折线散点图
        c.xlLine -- 折线图
        '''
        if sht:
            self.sheet = sht
        else:
            excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
            self.sheet = excel.ActiveSheet

        try:
            self.sheet.ChartObjects(chart_name).Delete()
        except:
            pass
        finally:
            self.chart_obj = self.sheet.Shapes.AddChart(chart_type,
                                                        chart_size[0],
                                                        chart_size[1],
                                                        chart_size[2],
                                                        chart_size[3])
            self.chart_obj.Name = chart_name
        self.chart = self.chart_obj.Chart

    def data_rng(self, rng=None, chart_plotby=2):
        '''

        c.xlColumns -- 2
        c.xlRows --1

        :param rng:
        :param ChartPlotBy: c.xlRows  or  c.xlColumns
        :return:
        '''
        if rng:
            rng = rng
        else:
            rng = self.sheet.UsedRange

        self.chart.SetSourceData(Source=rng, PlotBy=chart_plotby)

    def chartstyle_num(self, style_num=233):
        self.chart.ChartStyle = style_num

    def no_gridline(self, bool=False):
        self.chart.Axes(c.xlValue).HasMajorGridlines = bool

    def title(self, text):
        self.chart.HasTitle = True
        self.chart.ChartTitle.Text = text
        title_obj = self.chart.ChartTitle
        return title_obj

    def series(self, s_num):
        series_obj = self.chart.FullSeriesCollection(s_num)
        return series_obj

    def series_len(self):
        s_len = len(self.chart.FullSeriesCollection())
        return s_len

    def point(self, s_num, p_num):
        series_obj = self.series(s_num)
        point_obj = series_obj.Points(p_num)
        return point_obj

    def point_len(self, s_num):
        series_obj = self.series(s_num)
        point_len = len(series_obj.Points())
        return point_len

    def group_obj(self, g_num, which=None):
        obj = self.chart.ChartGroups(g_num)
        if which == 'down' or which == 1:
            obj.HasUpDownBars = True
            obj = obj.DownBars
        elif which == 'up' or which == 0:
            obj.HasUpDownBars = True
            obj = obj.UpBars
        return obj

    def data_table(self):
        self.chart.HasDataTable = True
        datetable_obj = self.chart.DataTable
        return datetable_obj

    def legend(self):
        self.chart.HasLegend = True
        legend_obj = self.chart.Legend
        return legend_obj

    def ticklabel(self, value=2):
        '''
        :param value: c.xlValue -- 2 Y轴  c.xlCategory -- 1 X轴
        :return:
        '''
        if value == 2:
            ticklabel_obj = self.chart.Axes(c.xlValue).TickLabels
        elif value == 1:
            ticklabel_obj = self.chart.Axes(c.xlCategory).TickLabels
        return ticklabel_obj

    def select(self):
        self.chart_obj.Select()

    def chart_obj(self):
        return self.chart_obj

    def point_has_datalabel(self, point_obj, bool=True):
        if bool:
            point_obj.ApplyDataLabels()
        else:
            point_obj.DataLabel.Delete()

    def series_style(self, series_obj, style_dict):
        if 'smooth' in style_dict:
            series_obj.Smooth = style_dict['smooth']
        if 'axis' in style_dict:
            series_obj.AxisGroup = style_dict['axis']
        if 'chart_type' in style_dict:
            series_obj.ChartType = style_dict['chart_type']

    def line_style(self, obj, style_dict):
        if 'color' in style_dict:
            obj.Format.Line.ForeColor.RGB = self.rgbToInt(style_dict['color'])
        if 'wieght' in style_dict:
            obj.Format.Line.Weight = style_dict['wieght']
        if 'visible' in style_dict:
            # c.msoTrue -- -1  c.msoFalse -- 0
            obj.Format.Line.Visible = style_dict['visible']
        if 'fillcolor' in style_dict:
            obj.Format.Fill.ForeColor.RGB = self.rgbToInt(style_dict['fillcolor'])

    def position(self, obj, pos):
        if isinstance(pos, (list, tuple)):
            obj.Left = pos[0]
            obj.Top = pos[1]
        elif isinstance(pos, int):
            obj.Position = pos

    def font_style(self, obj, style_dict):
        font = obj.Font
        if 'name' in style_dict:
            font.Name = style_dict['name']
        if 'size' in style_dict:
            font.Size = style_dict['size']
        if 'bold' in style_dict:
            font.Bold = style_dict['bold']
        if 'color' in style_dict:
            font.Color = self.rgbToInt(style_dict['color'])

    def rgbToInt(self, rgb):
        if isinstance(rgb, (list, tuple)):
            colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
        elif isinstance(rgb, int):
            colorInt = rgb
        return colorInt


def chart_trend(data_rng, chart_title, chart_name, series_lines_color_list):
    chart_size = (0, 170, 1200, 500)  # 大小
    chart_type = c.xlLine  # 类型

    chart = ExcelChart(chart_name=chart_name, chart_size=chart_size, chart_type=chart_type)

    chart_plotby = c.xlRows
    chart.data_rng(rng=data_rng, chart_plotby=chart_plotby)

    no_gridline = True  # 无网格线
    chart.no_gridline(no_gridline)

    chart_style_num = 233  # 格式
    chart.chartstyle_num(chart_style_num)

    font = {'name': 'Microsoft YaHei UI'}

    # 标题
    title_font = font.copy()
    title_font['size'] = 24
    title_obj = chart.title(chart_title)
    chart.font_style(obj=title_obj, style_dict=title_font)

    # 除标题外 其余字体
    other_font = font.copy()
    other_font['size'] = 12

    # 图例
    legend_obj = chart.legend()
    legend_font = other_font.copy()
    chart.font_style(obj=legend_obj, style_dict=legend_font)
    legend_pos = c.xlLegendPositionTop
    chart.position(obj=legend_obj, pos=legend_pos)

    # Y轴
    ticklabel_obj = chart.ticklabel()
    ticklabel_font = other_font.copy()
    chart.font_style(obj=ticklabel_obj, style_dict=ticklabel_font)

    # 数据表
    datatable_obj = chart.data_table()
    datatable_font = other_font.copy()
    chart.font_style(obj=datatable_obj, style_dict=datatable_font)

    # 标签
    series_style = {'smooth': True}  # 标签平滑

    series_len = chart.series_len()
    position_list = list(range(1, series_len + 1))
    series_lines_color_list = series_lines_color_list[::-1]
    position_list = position_list[::-1]
    line_color_dict = {}
    for n, i in enumerate(position_list):
        series_obj = chart.series(i)
        chart.series_style(series_obj=series_obj, style_dict=series_style)
        line_color_dict['color'] = series_lines_color_list[n]
        chart.line_style(obj=series_obj, style_dict=line_color_dict)

    chart.select()
    return chart


def chart_component(data_rng, chart_title, chart_name):
    chart = chart_trend(data_rng, chart_title, chart_name,
                        series_lines_color_list=component_RGB)
    return chart


def chart_4monthes_trend(data_rng, chart_title, chart_name):
    chart = chart_trend(data_rng, chart_title, chart_name,
                        series_lines_color_list=RGBs)
    p_pos = [630, 80]
    s1_num = chart.series_len()
    s1_label_style = {'name': 'Microsoft YaHei UI', 'size': 15, 'bold': True, 'color': (255, 255, 255)}
    point_obj = chart.point(s_num=s1_num, p_num=15)
    chart.point_has_datalabel(point_obj)
    p_datalabel_obj = point_obj.DataLabel
    chart.font_style(obj=p_datalabel_obj, style_dict=s1_label_style)
    chart.position(obj=p_datalabel_obj, pos=p_pos)

    s2_num = s1_num - 1
    s2_label_style = {'name': 'Microsoft YaHei UI', 'size': 12, 'bold': True, 'color': (255, 0, 0)}
    series_label_pos = c.xlLabelPositionAbove
    s2_p_len = chart.point_len(s2_num)
    for _ in range(1, s2_p_len + 1):
        s2_p_obj = chart.point(s_num=s2_num, p_num=_)
        chart.point_has_datalabel(s2_p_obj)
        s2_p_datatable_obj = s2_p_obj.DataLabel
        chart.position(obj=s2_p_datatable_obj, pos=series_label_pos)
        chart.font_style(obj=s2_p_datatable_obj, style_dict=s2_label_style)
    chart.select()
    return chart


def chart_common(data_rng, chart_title, chart_name, chart_type):
    chart_size = (0, 170, 900, 400)  # 大小

    chart = ExcelChart(chart_name=chart_name, chart_size=chart_size, chart_type=chart_type)

    chart_plotby = c.xlColumns
    chart.data_rng(rng=data_rng, chart_plotby=chart_plotby)

    no_gridline = False  # 无网格线
    chart.no_gridline(no_gridline)

    font = {'name': 'Microsoft YaHei UI'}

    # 标题
    title_font = font.copy()
    # title_font['size'] = 24
    title_obj = chart.title(chart_title)
    chart.font_style(obj=title_obj, style_dict=title_font)

    # 除标题外 其余字体
    other_font = font.copy()
    other_font['size'] = 12

    # 图例
    legend_obj = chart.legend()
    legend_font = other_font.copy()
    chart.font_style(obj=legend_obj, style_dict=legend_font)
    legend_pos = c.xlLegendPositionTop
    chart.position(obj=legend_obj, pos=legend_pos)

    # X轴 和 Y轴
    ticklabel_objs = [chart.ticklabel(value=1), chart.ticklabel(value=2)]
    ticklabel_font = other_font.copy()
    ticklabel_font['size'] = 10
    for _ in ticklabel_objs:
        chart.font_style(obj=_, style_dict=ticklabel_font)

    chart.select()
    return chart


def chart_vs_last_month(data_rng, chart_title, chart_name):
    chart_type = c.xlColumnClustered  # 类型
    chart = chart_common(data_rng, chart_title, chart_name, chart_type=chart_type)
    # 标签
    xlColumn_label_font = {'name': 'Microsoft YaHei UI', 'size': 14}
    for _ in range(1, 3):
        p_len = chart.point_len(s_num=_)
        for i in range(1, p_len + 1):
            p_obj = chart.point(s_num=_, p_num=i)
            chart.point_has_datalabel(point_obj=p_obj)
            p_datatable_obj = p_obj.DataLabel
            chart.font_style(obj=p_datatable_obj, style_dict=xlColumn_label_font)

    xlLine_label_pos = 0
    font_color = {3: 7434613, 4: 49407}
    xlLine_label_font = {'name': 'Microsoft YaHei UI', 'size': 15, 'bold': True}
    xlLine_label_line = {'color': (91, 155, 213), 'visible': -1, 'weight': 1.1}
    # c.msoTrue -- -1
    series_style = {'chart_type': c.xlLine}  # 3 / 4 折线
    for _ in range(3, 5):
        series_obj = chart.series(s_num=_)
        chart.series_style(series_obj=series_obj, style_dict=series_style)
        p_len = chart.point_len(s_num=_)
        p_obj = chart.point(s_num=_, p_num=p_len)
        chart.point_has_datalabel(point_obj=p_obj)
        p_datatable_obj = p_obj.DataLabel
        chart.position(obj=p_datatable_obj, pos=xlLine_label_pos)
        xlLine_label_font['color'] = font_color[_]
        chart.font_style(obj=p_datatable_obj, style_dict=xlLine_label_font)
        chart.line_style(obj=p_datatable_obj, style_dict=xlLine_label_line)

    chart.select()
    return chart


def chart_complete_rate_time_rate(data_rng, chart_title, chart_name):
    chart_type = c.xlLineMarkers  # 类型
    chart = chart_common(data_rng, chart_title, chart_name, chart_type=chart_type)

    # 跌涨柱颜色
    downbar_obj = chart.group_obj(g_num=1, which=1)
    downbar_dict = {'fillcolor': (157, 11, 11)}
    upbar_obj = chart.group_obj(g_num=1, which=0)
    upbar_dict = {'fillcolor': (0, 121, 68)}
    chart.line_style(obj=downbar_obj, style_dict=downbar_dict)
    chart.line_style(obj=upbar_obj, style_dict=upbar_dict)

    # label_1
    p_len = chart.point_len(s_num=1)
    s1_datalabel_pos = 0  # xlLabelPositionAbove
    s1_line = {'visible': 0, 'fillcolor': (240, 240, 240)}  # c.msoFalse = 0
    s1_font = {'name': 'Microsoft YaHei UI', 'size': 15, 'bold': True, 'color': (91, 155, 213)}
    for _ in range(1, p_len + 1):
        p_obj = chart.point(s_num=1, p_num=_)
        chart.point_has_datalabel(point_obj=p_obj, bool=True)
        p_label_obj = p_obj.DataLabel
        chart.position(obj=p_label_obj, pos=s1_datalabel_pos)
        chart.line_style(obj=p_label_obj, style_dict=s1_line)
        chart.font_style(obj=p_label_obj, style_dict=s1_font)

    # 横轴刻度 在外部调整
    chart.select()
    return chart

def chart_group_leader_rate(data_rng, chart_title, chart_name):
    chart_type = c.xlColumnStacked  # 类型
    chart = chart_common(data_rng, chart_title, chart_name, chart_type=chart_type)
    # 标签
    xlLine_label_pos = 0
    xlLine_label_font = {'name': 'Microsoft YaHei UI', 'size': 12, 'bold': True,'color':(19, 51, 76)}
    xlLine_label_line = {'color': (91, 155, 213), 'visible': -1, 'weight': 1.1,'fillcolor':(246, 246, 233)}
    # c.msoTrue -- -1
    series_style = {'chart_type': c.xlLine,'axis':2}  # 3 / 4 折线
    series_obj = chart.series(s_num=3)
    chart.series_style(series_obj=series_obj, style_dict=series_style)
    p_len = chart.point_len(s_num=3)
    for _ in range(1,p_len+1):
        p_obj = chart.point(s_num=3, p_num=_)
        chart.point_has_datalabel(point_obj=p_obj)
        p_datatable_obj = p_obj.DataLabel
        chart.position(obj=p_datatable_obj, pos=xlLine_label_pos)
        chart.font_style(obj=p_datatable_obj, style_dict=xlLine_label_font)
        chart.line_style(obj=p_datatable_obj, style_dict=xlLine_label_line)

    chart.select()
    return chart

#todo
def chart_team_rank(data_rng, chart_title, chart_name):
    chart_type = c.xlColumnStacked  # 类型
    chart = chart_common(data_rng, chart_title, chart_name, chart_type=chart_type)
    # 标签
    xlLine_label_pos = 0
    xlLine_label_font = {'name': 'Microsoft YaHei UI', 'size': 12, 'bold': True,'color':(19, 51, 76)}
    xlLine1_label_line = {'color': (91, 155, 213), 'visible': -1, 'weight': 1.1,'fillcolor':(246, 246, 233)}
    xlLine2_label_line = {'color': (91, 155, 213), 'visible': -1, 'weight': 1.1,'fillcolor':(246, 246, 233)}
    # c.msoTrue -- -1
    series_style = {'chart_type': c.xlLine,'axis':2}  # 3 / 4 折线
    series_obj = chart.series(s_num=3)
    chart.series_style(series_obj=series_obj, style_dict=series_style)
    p_len = chart.point_len(s_num=3)
    for _ in range(1,p_len+1):
        p_obj = chart.point(s_num=3, p_num=_)
        chart.point_has_datalabel(point_obj=p_obj)
        p_datatable_obj = p_obj.DataLabel
        chart.position(obj=p_datatable_obj, pos=xlLine_label_pos)
        chart.font_style(obj=p_datatable_obj, style_dict=xlLine_label_font)
        chart.line_style(obj=p_datatable_obj, style_dict=xlLine_label_line)

    chart.select()
    return chart
