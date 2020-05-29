#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/12 9:37:14
# @Author  : HouWk
# @Site    : 
# @File    : cls_excel.py
# @Software: PyCharm
import sys
from copy import deepcopy

import win32com.client
from win32com.client import constants as c  # 旨在直接使用VBA常数
import os
import numpy as np
from cls_data_dataframe import DataAsDF
from cls_date import MyDate
from cls_sqlserver import ReportDataAsDf
from fun_date import get_str_date
from use_os import create_folder_date


class Excel:
    def __init__(self):
        self.excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        self.excel.DisplayAlerts = 0
        self.excel.SheetsInNewWorkbook = 1
        # 全局字体
        self.font_name = 'Microsoft YaHei UI'
        self.big_size = 24
        self.medium_size = 12
        self.small_size = 10

    def screen_updating(self, bol=True):
        self.excel.ScreenUpdating = bol
        self.excel.Visible = bol

    def file_rename(self, path, file_name):
        full_name = os.path.join(path, file_name)
        only_file_name, extention_name = os.path.splitext(file_name)
        i = 1
        while os.path.isfile(full_name):
            new_file_name = '{}{}{}'.format(only_file_name, '(%d)' % i, extention_name)
            full_name = os.path.join(path, new_file_name)
            i += 1
        return full_name

    def workbook_save(self, wookbook_obj, name, path):
        MyType = '.xlsx'
        if not name.endswith(MyType):
            name = name + MyType
        full_name = self.file_rename(path, name)
        wookbook_obj.SaveAs(full_name)
        wookbook_obj.Save()
        return wookbook_obj

    # 数据
    def write_data(self, start_cel, data):
        try:
            if isinstance(data, (tuple, list)):
                shape = (np.array(data)).shape
                for i in range(shape[0]):
                    if len(shape) == 2:
                        for j in range(shape[1]):
                            start_cel.GetOffset(i, j).Value = data[i][j]
                    else:
                        start_cel.GetOffset(0, i).Value = data[i]
        except Exception as err:
            print('err:', err)

    def font_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'size':
                obj.Size = v
            elif k == 'color':
                obj.Color = self.rgbToInt(v)
            elif k == 'bold':
                obj.Bold = v
            elif k == 'name':
                obj.Name = v
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    def rgbToInt(self, rgb):
        if isinstance(rgb, (list, tuple)):
            colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
        elif isinstance(rgb, int):
            colorInt = rgb
        return colorInt

    def autofit(self, rng, columnlist):
        col_inds = []
        # 将自动调整列宽的列下标写进列表
        if isinstance(columnlist, list):
            col_inds = columnlist
        elif isinstance(columnlist, int):
            col_inds.append(columnlist)
        for col_ind in col_inds:
            rng.Columns(col_ind).AutoFit()

    def bold(self, rng, column_list, tag='总计'):
        rc = rng.Rows.Count
        for i in range(1, rc + 1):
            for j in column_list:
                if rng.Cells(i, j).Value:
                    if rng.Cells(i, j).Value.find(tag) >= 0:
                        rng.Rows(i).Font.Bold = True
                        continue

    def merge(self, sht_obj,rng, column_list):
        rc = rng.Rows.Count
        for j in column_list:
            for i in range(rc, 1, -1):
                if j == 1:
                    if not (rng.Cells(i, j).Value and rng.Cells(i - 1, j).Value):
                        sht_obj.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()
                    elif rng.Cells(i, j).Value == rng.Cells(i - 1, j).Value:
                        sht_obj.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()
                elif j > 1:
                    if sht_obj.Range(rng.Cells(i, j - 1), rng.Cells(i - 1, j - 1)).MergeCells:
                        if (not rng.Cells(i, j).Value) and (not rng.Cells(i - 1, j).Value):
                            sht_obj.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()
                        elif rng.Cells(i, j).Value == rng.Cells(i - 1, j).Value:
                            sht_obj.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()

    def xl3Triangles(self, rng, columnlist):
        col_inds = []
        if isinstance(columnlist, list):
            col_inds = columnlist
        elif isinstance(columnlist, int):
            col_inds.append(columnlist)
        for col_ind in col_inds:
            rng_col = rng.Columns(col_ind)
            rng_col.FormatConditions.AddIconSetCondition()
            rng_col_fc1 = rng.FormatConditions(1)
            rng_col_fc1.IconSet = self.excel.ActiveWorkbook.IconSets(c.xl3Triangles)
            rng_col_fc1_ic2 = rng_col_fc1.IconCriteria(2)
            rng_col_fc1_ic3 = rng_col_fc1.IconCriteria(3)
            rng_col_fc1_ic2.Type = c.xlConditionValueNumber
            rng_col_fc1_ic2.Operator = 7
            rng_col_fc1_ic2.Value = 0
            rng_col_fc1_ic3.Type = c.xlConditionValueNumber
            rng_col_fc1_ic3.Operator = 5
            rng_col_fc1_ic3.Value = 0

    def xlConditionValueNumber(self, rng, columnlist):
        col_inds = []
        if isinstance(columnlist, list):
            col_inds = columnlist
        elif isinstance(columnlist, int):
            col_inds.append(columnlist)
        for col_ind in col_inds:
            rng_col = rng.Columns(col_ind)
            rng_col.Style = "Percent"
            rng_col.FormatConditions.AddDatabar()
            rng_col_fc1 = rng.FormatConditions(1)
            rng_col_fc1.MinPoint.Modify(newtype=c.xlConditionValueNumber, newvalue=0)
            rng_col_fc1.MaxPoint.Modify(newtype=c.xlConditionValueNumber, newvalue=1)
            rng_col_fc1.BarColor.Color = 13012579
            rng_col_fc1.BarColor.TintAndShade = 0

    # 图
    def position(self, obj, pos):
        if isinstance(pos, (list, tuple)):
            obj.Left = pos[0]
            obj.Top = pos[1]
        elif isinstance(pos, int):
            obj.Position = pos


    def obj_list_series(self,series_all_obj_list,series_num):
        if series_num == 'all':
            series_real_obj_list = series_all_obj_list
        elif isinstance(series_num, int):
            series_real_obj_list = []
            series_real_obj_list.append(series_all_obj_list[series_num])
        elif isinstance(series_num, (list, tuple)):
            series_real_obj_list = []
            for _ in series_num:
                series_real_obj_list.append(series_all_obj_list[_])
        else:
            raise AttributeError('Not Existed {}.{}'.format('chart', series_num))

        return series_real_obj_list


    #数据条
    def obj_updownbars(self, chart_obj,chartgroup_num=1):
        #返回 [db,ub]
        obj = chart_obj.ChartGroups(chartgroup_num)
        obj.HasUpDownBars = True
        downb_obj = obj.DownBars
        up_obj = obj.UpBars
        return [downb_obj,up_obj]

    # 线条
    def line_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'forecolor':
                forecolor_obj = obj.ForeColor
                self.forecolor_style(forecolor_obj,**v)
            elif k == 'weight':
                obj.Weight = v
            elif k == 'visible':
                # c.msoTrue -- -1  c.msoFalse -- 0
                obj.Visible = v
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    # 数据区域 标题
    def range_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'value':
                obj.Value = v
            elif k == 'font':
                title_font_obj = obj.Font
                self.font_style(title_font_obj, **v)
            elif k == 'merge':
                if v == True:
                    try:
                        obj.Merge()
                    except:
                        pass
            elif k == 'horizontalalignment':
                obj.HorizontalAlignment = v
            elif k == 'verticalalignment':
                obj.VerticalAlignment = v
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    # 行格式
    def row_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'rowheight':
                obj.RowHeight = v
            elif k=='wraptext':
                obj.WrapText = v
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    # 标题
    def title_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'text':
                obj.Text = v
            elif k == 'font':
                title_font_obj = obj.Font
                self.font_style(title_font_obj, **v)
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    # 图例
    def legend_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'position':
                self.position(obj=obj, pos=v)
            elif k == 'font':
                title_font_obj = obj.Font
                self.font_style(title_font_obj, **v)
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    # 标签
    def datalabel_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'position':
                self.position(obj=obj, pos=v)
            elif k == 'font':
                font_obj = obj.Font
                self.font_style(font_obj, **v)
            elif k == 'format':
                obj = obj.Format
                self.format_style(obj, **v)
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    # 坐标轴
    def ticklabel_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'font':
                ticklabel_font_obj = obj.Font
                self.font_style(ticklabel_font_obj, **v)
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    # 坐标轴
    def datatable_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'font':
                datatable_font_obj = obj.Font
                self.font_style(datatable_font_obj, **v)
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    # 序列
    def series_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'smooth':
                obj.Smooth = v
            elif k == 'axisgroup':
                obj.AxisGroup = v
            elif k == 'charttype':
                # c.msoTrue -- -1  c.msoFalse -- 0
                obj.ChartType = v
            elif k == 'format':
                obj = obj.Format
                self.format_style(obj, **v)
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    def series_linecolors_list(self,obj_list,**linecolors):
        '''
        linecolors:{
            'reverse':True
            'colors':[]
        }
        :param obj_list:
        :param linecolors_list:
        :return:
        '''
        color_list = linecolors['colors']
        len_color_list = len(color_list)
        if ('reverse' in linecolors) and (linecolors['reverse'] == True):
            obj_list = obj_list[::-1]
            color_list = color_list[::-1]
        for i,obj in enumerate(obj_list):
            j = (len_color_list-1) if i >= len_color_list else i
            obj.Format.Line.ForeColor.RGB = self.rgbToInt(color_list[j])

    #数据点
    def point_style(self,obj,**kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'datalabels':
                datalabel_obj = obj.DataLabel
                self.datalabel_style(datalabel_obj, **v)
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name, k))


    def updownbars_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'format':
                # obj.Format.Line.ForeColor.RGB = self.rgbToInt(style_dict['color'])
                obj = obj.Format
                self.format_style(obj,**v)
                # self.fill_style(obj, **v)
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    # 填充
    def fill_style(self, obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'forecolor':
                obj = obj.ForeColor
                self.forecolor_style(obj,**v)
                # obj.ForeColor.RGB = self.rgbToInt(v)  # 前景色
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    def format_style(self,obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'fill':
                fill_obj = obj.Fill
                self.fill_style(obj=fill_obj,**v)
            elif k=='line':
                line_obj = obj.Line
                self.line_style(obj=line_obj, **v)
            elif k == 'textframe2':
                TextFrame2_obj = obj.TextFrame2
                self.textframe2_style(obj=TextFrame2_obj, **v)
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    def textframe2_style(self,obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'orientation':
                obj.Orientation = v
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))



    def forecolor_style(self,obj, **kwargs):
        for k, v in kwargs.items():
            k = k.lower()
            if k == 'rgb':
                obj_rgb = obj.RGB
                obj.RGB =self.rgbToInt(v)
            else:
                raise AttributeError('Not Existed {}.{}'.format(sys._getframe().f_code.co_name,k))

    # 应用
    def common_sheet(self, wb_obj, sht_name, tab_color, df, title_v=None,
                     subtitles_v=None, shadow=False, gridlines=False):
        sheets_count = wb_obj.Worksheets.Count
        # 新增表
        wb_obj.Sheets.Add(After=wb_obj.Sheets(sheets_count)).Name = sht_name
        sht_obj = wb_obj.ActiveSheet
        # 无网格线
        if not gridlines:
            self.excel.ActiveWindow.DisplayGridlines = False
        # 表 标签颜色
        sht_tab_colors = {'R': 7697919, 'DR': 255, 'G': 11854022, 'DG': 2315831}
        if isinstance(tab_color, str):
            sht_tab_colors = sht_tab_colors[tab_color]
        sht_obj.Tab.Color = sht_tab_colors
        # df数据
        col_names = list(df.columns)
        result_value = list(df.values)
        cs = len(col_names)
        r_index = 1

        # 区域
        # 标题
        if title_v:
            cel_title = sht_obj.Range("a" + str(r_index))

            row_obj = sht_obj.Rows(r_index)
            row_dict = {'RowHeight': 30}
            self.row_style(row_obj, **row_dict)

            r_index += 1
            title_rg = sht_obj.Range(cel_title, cel_title.GetOffset(0, cs - 1))

            r_title = {
                'value':title_v,
                'font': {'name': self.font_name,
                         'bold': True,
                         'color': 7884319,
                         'size': 16
                         },
                'merge': True,
                'HorizontalAlignment': c.xlCenter,
                'VerticalAlignment': c.xlCenter
            }

            self.range_style(title_rg, **r_title)
        # 副标题
        if subtitles_v:
            cel_subtitle1 = sht_obj.Range("a" + str(r_index))
            r_index += 1
            cel_subtitle2 = cel_subtitle1.GetOffset(0, cs - 1)

            subtitle_rg1 = sht_obj.Range(cel_subtitle1, cel_subtitle2.GetOffset(0, -1))
            r1_subtitle = {
                'value':'统计时间：',
                'merge':True,
                'HorizontalAlignment':c.xlRight
            }
            self.range_style(subtitle_rg1,**r1_subtitle)

            subtitle_rg2 = cel_subtitle2
            r2_subtitle = {
                'value': subtitles_v,
                'HorizontalAlignment': c.xlLeft
            }
            self.range_style(subtitle_rg2, **r2_subtitle)

            subtitle_rg = sht_obj.Range(cel_subtitle1, cel_subtitle2)
            r_subtitle = {
                'font': {'name': self.font_name,
                         'size': 10
                         },
            }
            self.range_style(subtitle_rg, **r_subtitle)
        # 列
        r_cols = {
            'font': {'name': self.font_name,
                     'size': 10,
                     'bold': True,
                     'color': 16777215
                     },
            'HorizontalAlignment': c.xlCenter,
            'VerticalAlignment': c.xlCenter
        }
        cel_cols = sht_obj.Range("a" + str(r_index))
        self.write_data(cel_cols, col_names)
        r_index += 1
        cols_rg = sht_obj.Range(cel_cols, cel_cols.GetOffset(0, cs - 1))
        self.range_style(cols_rg, **r_cols)
        cols_rg.Cells.EntireRow.AutoFit()
        cols_rg.Cells.EntireColumn.AutoFit()
        cols_rg.Borders.LineStyle = c.xlContinuous
        cols_rg.Borders.Weight = c.xlThin
        cols_rg.Borders.ThemeColor = 1
        cols_rg.Borders.TintAndShade = -0.14996795556505
        cols_rg.Interior.Color = 7884319
        # 数据区域
        cel_data = sht_obj.Range("a" + str(r_index))
        self.write_data(cel_data, result_value)  # 数据
        rg = sht_obj.UsedRange
        rs = rg.Rows.Count
        data_rg = sht_obj.Range(cel_data, sht_obj.Cells(rs, cs))
        data_rg.HorizontalAlignment = c.xlCenter
        data_rg.VerticalAlignment = c.xlCenter
        data_rg.Borders.LineStyle = c.xlContinuous
        data_rg.Borders.Weight = c.xlThin
        data_rg_dict = {
            'font':{'name':self.font_name,
                   'size':10},
            }
        self.range_style(data_rg, **data_rg_dict)
        data_rg.Borders.ThemeColor = 1
        data_rg.Borders.TintAndShade = -0.14996795556505
        data_rg.Borders(c.xlEdgeBottom).Color = 5
        data_rg.Borders(c.xlEdgeBottom).TintAndShade = -0.499984740745262
        data_rg.Cells.EntireColumn.AutoFit()
        data_rg.Cells.EntireRow.AutoFit()
        if shadow:
            if rs > 2:
                for i in range(1, rs, 2):
                    data_rg.Rows(i).Interior.Color = 15921906
        rg.Select()
        return sht_obj, rs, cs

    def common_chart(self, sht_obj, chart_data_rng_obj, chart_name, chart_type, chart_plotby, chart_style_num=None,
                     chart_size=(0, 170, 900, 400),
                     gridline=False, chart_title=None, legend=None, ticklabel=None,
                     datatable=None, series=None, point=None, updownbars=None):
        '''
        :param sht_obj:
        :param chart_data_rng_obj:
        :param chart_name:
        :param chart_type: c.xlColumnStacked -- 柱形堆积图
                        c.xlColumnClustered -- 柱形图
                        c.xlLineMarkers -- 折线散点图
                        c.xlLine -- 折线图
        :param chart_style_num:
        :param chart_size:
        :param chart_plotby: 行 c.xlColumns 或 列 c.xlRows
        :param gridline:
        :param chart_title:
        :param legend:
        :param ticklabel:
        :param datatable:
        :param series:['all':{'smooth':,
                            'axisgroup':,
                            'charttype':,
                            'line':{
                                    'wieght':,
                                    'visible':,
                                    'color_list_dict':{'reverse':,
                                                        'color_list':
                                                        }
                                    }
                            },
                        ]
        :param point:[{'s_num':'all','p_num':'all',
                    'style':{
                            'font':font_dict,
                            'pos':pos_dict,
                            'line':{'weight':,
                                    'visible': ,
                                    'color': ,
                                    }
                            }
                    },
                ]
        :return:
        '''
        chart_obj = sht_obj.Shapes.AddChart(chart_type,
                                            chart_size[0],
                                            chart_size[1],
                                            chart_size[2],
                                            chart_size[3])
        chart_obj.Name = chart_name
        chart = chart_obj.Chart

        # 数据区域和图类型
        chart.SetSourceData(Source=chart_data_rng_obj, PlotBy=chart_plotby)

        # 无网格线
        chart.Axes(c.xlValue).HasMajorGridlines = gridline

        # 图整体样式
        if chart_style_num:
            chart.ChartStyle = chart_style_num

        # 标题
        if chart_title:
            chart.HasTitle = True
            title_obj = chart.ChartTitle
            self.title_style(title_obj, **chart_title)

        # 图例
        if legend:
            chart.HasLegend = True
            legend_obj = chart.Legend
            self.legend_style(legend_obj, **legend)

        # X轴 和 Y轴
        if ticklabel:
            ticklabel_objs = [chart.Axes(c.xlValue).TickLabels, chart.Axes(c.xlCategory).TickLabels]
            for _ in ticklabel_objs:
                self.ticklabel_style(obj=_, **ticklabel)

        # 数据表
        if datatable:
            chart.HasDataTable = True
            datetable_obj = chart.DataTable
            self.datatable_style(obj=datetable_obj, **datatable)

        if series:
            series_all_obj_list = [_ for _ in chart.FullSeriesCollection()]

            for serie_dict in series:
                for k, v in serie_dict.items():
                    # 确定series元素
                    obj_list_series = self.obj_list_series(series_all_obj_list,k)

                    for k1,v1 in v.items():
                        k1 = k1.lower()
                        if k1 == 'linecolors':
                            self.series_linecolors_list(obj_list_series,**v1)
                        else:
                            v2 = {k1:v1}
                            for series in obj_list_series:
                                self.series_style(series,**v2)

        if point:
            p_obj_list = [[j  for j in i.Points()] for i in chart.FullSeriesCollection()]
            p_obj_arr = np.array(p_obj_list)

            for point_dict in point:
                for k, v in point_dict.items():
                    # 确定series元素
                    p_obj_real_list = eval('p_obj_arr'+k)

                    if not isinstance(p_obj_real_list,np.ndarray):
                        p_obj_real_list.HasDataLabel = True
                        self.point_style(p_obj_real_list, **v)
                    else:
                        for ind,p_obj in np.ndenumerate(p_obj_real_list):
                            p_obj.HasDataLabel = True
                            self.point_style(p_obj,**v)


        if updownbars:
            '''
            [
            {'upbars':
                'format':{
                    'fill':{
                        'ForeColor':{
                            'RGB':(157, 11, 11)
                            }
                        }
                    }
                },
            {'downbars':
                'format':{
                    'fill':{
                        'ForeColor':{
                            'RGB':(157, 11, 11)
                            }
                        }
                    }
                }
            ]
            '''
            downbar_obj,upbar_obj = self.obj_updownbars(chart)
            for bars in updownbars:
                for k,v in bars.items():
                    k = k.lower()
                    if k=='upbars':
                        obj = upbar_obj
                        self.updownbars_style(obj,**v)
                    elif k=='downbars':
                        obj = downbar_obj
                        self.updownbars_style(obj,**v)
                    else:
                        raise AttributeError('Not Existed {}.{}'.format('chart', k))

        chart_obj.Select()
        return chart

    def subtitle(self, *args, format='%Y/%m/%d'):
        result = []
        for _ in args:
            dt = get_str_date(_, format=format)
            result.append(dt)
        subtitle_str = '-'.join(result)
        return subtitle_str



# if __name__ == '__main__':
    # a = Excel('2020/5/19').report_component()
    # a = Excel('2020/5/19').reports_day()