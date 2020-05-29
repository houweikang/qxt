#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/21 9:41:28
# @Author  : HouWk
# @Site    : 
# @File    : cls_reports.py
# @Software: PyCharm
import pandas as pd

from cls_excel import Excel
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


class Reports(Excel):
    def __init__(self, final_date):
        super().__init__()

        # 报表储存路径
        self.root_path = r'e:\python_Reports'

        # 终止日期
        self.final_date = final_date

        # 本月起始日期
        self.start_date = MyDate(final_date).get_date_Nmonthes_firstday(n=0)
        self.start_date = get_str_date(self.start_date)

        # 三个月前 起始日期
        self.start_date_date_3 = MyDate(final_date).get_date_Nmonthes_firstday(n=-3)  # 获取起始日期
        self.start_str_date_3 = get_str_date(self.start_date_date_3)  # 获取起始日期

        # 三个月前 起始日期
        self.start_date_date_6 = MyDate(final_date).get_date_Nmonthes_firstday(n=-5)  # 获取起始日期
        self.start_str_date_6 = get_str_date(self.start_date_date_6)  # 获取起始日期

        # 一个月前 起始日期
        self.start_str_date_1 = MyDate(final_date).get_date_Nmonthes_firstday(n=-1)  # 获取起始日期
        self.start_str_date_1 = get_str_date(self.start_str_date_1)  # 获取起始日期

        # 分量地区
        self.component_dqs = ['济南', '燕郊', '成都']
        # 分量图线条颜色
        self.RGBs = [(255, 255, 0), (0, 112, 192), (0, 176, 80), (255, 0, 0), (255, 255, 255)]  # 黄 蓝 绿 红 白
        self.component_RGB = self.RGBs[:4]

        # 其余报表地区
        self.common_dqs = ['保定', '济南']

        # 图大小
        self.size_long_chart = (0, 170, 1200, 500)
        self.size_chart = (0, 170, 900, 400)
        self.size_chart_datatable = (0, 170, 900, 500)

        self.size_chart_sixmonth = [(800, 0, 900, 400), (800, 401, 900, 400)]
        self.size_chart_sixmonth_group = [(800, 0, 900, 600), (800, 601, 900, 600)]

        # charttype 0-xlColumnClustered
        self.chart_type = [c.xlColumnClustered, ]

        # chartplotby 0-xlRows  1-xlColumns
        self.chart_plotby = [c.xlRows, c.xlColumns]

        # chart_style_num  0-233
        self.chart_style_num = [233, ]

        # datalabelposition 0、1 折线  2、3 柱形图 0-数据点上方 1-数据点下方 2-上边缘上方 3-上边缘下方
        self.datalabel_position = [c.xlLabelPositionAbove, c.xlLabelPositionBelow,
                                   c.xlLabelPositionOutsideEnd, c.xlLabelPositionInsideEnd]



    def chart_title(self, text):
        _dict = {
            'text': text,
            'font': {
                'name': self.font_name,
                'size': self.big_size
            }
        }
        return _dict

    def chart_legend(self):
        _dict = {
            'font': {
                'name': self.font_name,
                'size': self.medium_size
            },
            'position': c.xlLegendPositionTop
        }
        return _dict

    def chart_ticklabel(self):
        _dict = {
            'font': {
                'name': self.font_name,
                'size': self.medium_size
            },
        }
        return _dict

    def chart_datatable(self):
        _dict = {
            'font': {
                'name': self.font_name,
                'size': self.medium_size
            },
        }
        return _dict

    # 报表 分量
    def report_component(self):
        self.screen_updating(False)  # 关闭屏幕刷新
        wb_obj = self.excel.Workbooks.Add()  # 创建excel wb
        start_date = self.start_str_date_3  # 起始日期
        final_date = self.final_date  # 终止日期
        for dq in self.component_dqs:
            df = ReportDataAsDf(dq=dq, start_date=start_date, final_date=final_date).component_df()  # 数据帧
            # 表
            sheet_name = '%s近%d月日人均分量' % (dq, 4)

            # 区域标题
            r_title_v = sheet_name
            # 区域副标题
            r_subtitles_v = self.subtitle(start_date, final_date)
            # 区域赋值 并 设定格式
            sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color='R', df=df,
                                                title_v=r_title_v, subtitles_v=r_subtitles_v)
            # 图
            chart_size = (0, 170, 1200, 500)
            chart_type = c.xlLine
            chart_name = sheet_name
            chart_plotby = c.xlRows
            chart_style_num = 233
            gridline = False
            # 标题
            chart_title = '%s%s-近%d个月每日人均分量趋势' % ('QX', dq, 4)
            chart_title_dict = self.chart_title(chart_title)
            # 数据区域
            chart_data_rng = sht_obj.Range('a3', sht_obj.Cells(rs, cs))
            # 图例
            legend_dict = self.chart_legend()
            # 坐标轴
            ticklabel_dict = self.chart_ticklabel()
            # 数据表
            datatable_dict = self.chart_datatable()
            # 数据系列标签
            series_dict = [
                {'all': {
                    'smooth': True,
                    'linecolors': {
                        'colors': self.component_RGB,
                        'reverse': True
                    }
                }
                }
            ]

            self.common_chart(sht_obj=sht_obj, chart_data_rng_obj=chart_data_rng, chart_name=chart_name,
                              chart_type=chart_type, chart_style_num=chart_style_num, chart_size=chart_size,
                              chart_plotby=chart_plotby, gridline=gridline, chart_title=chart_title_dict,
                              legend=legend_dict,
                              ticklabel=ticklabel_dict, datatable=datatable_dict, series=series_dict)

        wb_obj.Sheets(1).Delete()
        wb_obj.Sheets(1).Select()
        # 保存文件
        path = create_folder_date(self.root_path, final_date)  # 创建目标文件夹
        str_date = get_str_date(final_date, '%Y%m%d')
        wb_name = '%s日报_人均分量' % (str_date)
        self.workbook_save(wookbook_obj=wb_obj, name=wb_name, path=path)
        self.screen_updating(True)  # 开启屏幕刷新

    # 日报 总
    def reports_day(self):
        for dq in self.common_dqs:
            # self.screen_updating(False)  # 关闭屏幕刷新

            self.screen_updating(True)  # 开启屏幕刷新

            wb_obj = self.excel.Workbooks.Add()  # 创建excel wb
            # 进群率和注册率
            group_data, team_data, region_data, colege_data = ReportDataAsDf(dq=dq, start_date=self.start_str_date_3,
                                     final_date=self.final_date).erollment_and_group_entry_rate_df()

            # 近6个月同期对比
            # sixmonth_group, sixmonth_team, sixmonth_colege, sixmonth_region = ReportDataAsDf(dq,
            #                                                                                  self.start_str_date_6,
            #                                                                                  self.final_date).team_six_month()
            tab_colors = ['R','DR','G','DG']
            # T
            # self.report_erollment_and_group_entry_rate(wb_obj=wb_obj,tab_color=tab_colors[0],df=region_data)
            # self.report_erollment_and_group_entry_rate(wb_obj=wb_obj,tab_color=tab_colors[0],df=colege_data)
            # self.report_erollment_and_group_entry_rate(wb_obj=wb_obj,tab_color=tab_colors[0],df=team_data)
            # self.report_4monthes_trend(dq=dq, wb_obj=wb_obj)
            # self.reports_vs_last_month(dq=dq, wb_obj=wb_obj)
            # self.reports_complete_rate_time_rate(dq=dq, wb_obj=wb_obj)
            # self.reports_group_leader_rate(dq=dq, wb_obj=wb_obj)
            # self.reports_team_day_evening(dq=dq, wb_obj=wb_obj)
            # self.reports_six_month(dq=dq, wb_obj=wb_obj, df=sixmonth_region, tab_color=tab_colors[0])
            # self.reports_six_month(dq=dq, wb_obj=wb_obj, df=sixmonth_colege, tab_color=tab_colors[0])
            # self.reports_six_month(dq=dq, wb_obj=wb_obj, df=sixmonth_team, tab_color=tab_colors[0])
            # self.reports_evening(dq=dq, wb_obj=wb_obj)

            # G
            # self.reports_group_peoplelist(dq=dq, wb_obj=wb_obj)
            # self.reports_complete_rate(dq=dq, wb_obj=wb_obj)
            # self.report_erollment_and_group_entry_rate(wb_obj=wb_obj,tab_color=tab_colors[2],df=group_data)


            # self.reports_six_month(dq=dq, wb_obj=wb_obj, df=sixmonth_group, tab_color=tab_colors[2])


            wb_obj.Sheets(1).Delete()
            wb_obj.Sheets(1).Select()

            # 保存文件
            path = create_folder_date(self.root_path, self.final_date)  # 创建目标文件夹
            str_date = get_str_date(self.final_date, '%Y%m%d')
            wb_name = '%s%s日报' % (str_date, dq)
            self.workbook_save(wookbook_obj=wb_obj, name=wb_name, path=path)
            # self.screen_updating(True)  # 开启屏幕刷新

    # 日报 - 注册率与进群率
    def report_erollment_and_group_entry_rate(self, wb_obj, tab_color,df):
        start_date = self.start_str_date_3
        final_date = self.final_date
        dep = df.columns[-6]
        sheet_name = '%s推广进群率与注册率统计' % dep
        title = sheet_name
        subtitles = self.subtitle(start_date, final_date)
        self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                          title_v=title, subtitles_v=subtitles)

    def chart_4monthes_trend(self, dep, sht_obj, chart_data_rng):
        # 图
        chart_name = '%s日创量趋势' % dep
        chart_title = '%s%s-近%d个月每日创量总业绩趋势' % ('QX', dep, 4)
        chart_plotby = c.xlRows
        # 标题
        chart_title_dict = self.chart_title(chart_title)
        # 图例
        legend_dict = self.chart_legend()
        # 坐标轴
        ticklabel_dict = self.chart_ticklabel()
        # 数据表
        datatable_dict = self.chart_datatable()
        # 系列
        series_dict = [
            {'all': {
                'smooth': True,
                'linecolors': {
                    'colors': self.RGBs,
                    'reverse': True
                }
            }
            }
        ]
        # 标签1
        s1_dict = {
            '[-1,15]': {
                'datalabels': {
                    'font': {
                        'name': self.font_name,
                        'size': self.medium_size,
                        'bold': True,
                        'color': (255, 255, 255)
                    },
                    'position': [630, 80]
                }
            }
        }

        # 标签2
        s2_dict = {
            '[-2,:]': {
                'datalabels': {
                    'font': {
                        'name': self.font_name,
                        'size': self.medium_size,
                        'bold': True,
                        'color': (255, 0, 0)
                    },
                    'position': c.xlLabelPositionAbove
                }
            }
        }
        p_list = [s1_dict, s2_dict]
        self.common_chart(sht_obj=sht_obj, chart_data_rng_obj=chart_data_rng, chart_name=chart_name,
                          chart_type=c.xlLine, chart_style_num=233, chart_size=(0, 170, 1200, 500),
                          chart_plotby=chart_plotby, gridline=False, chart_title=chart_title_dict,
                          legend=legend_dict, ticklabel=ticklabel_dict, datatable=datatable_dict,
                          series=series_dict, point=p_list)

    # 日报 - 近4个月趋势
    def report_4monthes_trend(self, dq, wb_obj):
        start_date = self.start_str_date_3
        final_date = self.final_date
        # 数据帧
        dfs_region, dfs_colege, dfs_team = ReportDataAsDf(dq=dq, start_date=start_date,
                                                          final_date=final_date).trend_4monthes_df()
        for df in dfs_region:
            dep = df.iloc[0, 0]
            sheet_name = '%s日创量趋势' % dep
            title = sheet_name
            subtitles = self.subtitle(start_date, final_date)
            sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color='R', df=df,
                                                title_v=title, subtitles_v=subtitles)
            chart_data_rng = sht_obj.Range(sht_obj.Range('b3'), sht_obj.Cells(rs, cs))
            self.chart_4monthes_trend(dep, sht_obj, chart_data_rng)

        for df in dfs_colege:
            dep = ''.join(list(df.iloc[0, :2]))
            sheet_name = '%s日创量趋势' % dep
            title = sheet_name
            subtitles = self.subtitle(start_date, final_date)
            sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color='R', df=df,
                                                title_v=title, subtitles_v=subtitles)
            chart_data_rng = sht_obj.Range(sht_obj.Range('c3'), sht_obj.Cells(rs, cs))
            self.chart_4monthes_trend(dep, sht_obj, chart_data_rng)

        for df in dfs_team:
            dep = ''.join(list(df.iloc[0, :3]))
            sheet_name = '%s日创量趋势' % dep
            title = sheet_name
            subtitles = self.subtitle(start_date, final_date)
            sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color='R', df=df,
                                                title_v=title, subtitles_v=subtitles)
            chart_data_rng = sht_obj.Range(sht_obj.Range('d3'), sht_obj.Cells(rs, cs))
            self.chart_4monthes_trend(dep, sht_obj, chart_data_rng)

    # 日报 -- 与上月同期业绩对比
    def reports_vs_last_month(self, dq, wb_obj):
        start_date = self.start_date
        final_date = self.final_date
        df = ReportDataAsDf(dq, start_date, final_date).vs_last_month_df()
        tab_color = 'R'

        sheet_name = '与上月同期业绩对比'
        title = sheet_name
        subtitles = self.subtitle(start_date, final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)

        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(final_date, format='%m.%d')
        chart_title = '%s-%s推广业绩与上月同期对比' % (title_date_start, title_date_final)

        # 图
        chart_name = sheet_name
        charttype = c.xlColumnClustered  # 类型
        chart_plotby = c.xlColumns
        chart_data_rng = sht_obj.Range(sht_obj.Range('c3'), sht_obj.Cells(rs, 7))
        # 标题
        chart_title_dict = self.chart_title(chart_title)
        # 图例
        legend_dict = self.chart_legend()
        # 坐标轴
        ticklabel_dict = self.chart_ticklabel()
        # 数据系列标签
        series_dict = [
            {(-1, -2): {'charttype': c.xlLine}}
        ]
        # point
        s1_s2_dict = {
            '[0:2,:]': {
                'datalabels': {
                    'font': {
                        'name': self.font_name,
                        'size': self.medium_size
                    },
                }
            }
        }

        # s2_dict = s1_dict.copy()
        # s2_dict['s_num'] = 1

        s3_dict = {
            '[-2,-1]': {
                'datalabels': {
                    'font': {
                        'name': self.font_name,
                        'size': 15,
                        'bold': True,
                        'color': 7434613
                    },
                    'position': 0,
                    'format': {'line': {
                        'ForeColor': {'RGB': (91, 155, 213)},
                        'visible': -1,
                        'weight': 1.1
                    }
                    }
                }
            }
        }

        s4_dict = {
            '[-1,-1]': {
                'datalabels': {
                    'font': {
                        'name': self.font_name,
                        'size': 15,
                        'bold': True,
                        'color': 49407
                    },
                    'position': 0,
                    'format': {'line': {
                        'ForeColor': {'RGB': (91, 155, 213)},
                        'visible': -1,
                        'weight': 1.1
                    }
                    }
                }
            }
        }
        # s4_dict['s_num'] = -1
        # s4_dict['style']['font']['color'] = 49407

        p_dict = [s1_s2_dict, s3_dict, s4_dict]

        self.common_chart(sht_obj=sht_obj, chart_data_rng_obj=chart_data_rng, chart_name=chart_name,
                          chart_type=charttype, chart_size=self.size_chart,
                          chart_plotby=chart_plotby, gridline=False, chart_title=chart_title_dict,
                          legend=legend_dict,
                          ticklabel=ticklabel_dict, series=series_dict, point=p_dict)

    # 日报 -- 完成率与时间消耗率
    def reports_complete_rate_time_rate(self, dq, wb_obj):
        start_date = self.start_date
        final_date = self.final_date
        df = ReportDataAsDf(dq, start_date, final_date).complete_rate_time_rate_df()
        tab_color = 'R'

        sheet_name = '完成率与时间消耗率'
        title = '{}各运营部{}'.format(dq, sheet_name)
        subtitles = self.subtitle(start_date, final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        # 图
        chart_name = sheet_name
        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(final_date, format='%m.%d')
        chart_title = '%s-%s完成率与时间消耗率涨跌图' % (title_date_start, title_date_final)
        chart_plotby = c.xlColumns
        charttype = c.xlLineMarkers
        chart_size = (0, 170, 900, 400)
        chart_rng = sht_obj.Range('c3', 'e%d' % rs)
        # 标题
        chart_title_dict = self.chart_title(chart_title)
        # 图例
        legend_dict = self.chart_legend()
        # 坐标轴
        ticklabel_dict = self.chart_ticklabel()
        # 系列
        series_dict = [
            {'all': {
                'format': {'line': {'visible': 0}}
            }
            }
        ]

        # 标签1
        s1_dict = {
            '[0,:]': {
                'datalabels': {
                    'font': {
                        'name': self.font_name,
                        'size': self.medium_size,
                        'bold': True,
                        'color': (91, 155, 213)
                    },
                    'position': 0,
                    'format': {
                        'fill': {
                            'forecolor': {
                                'rgb': (240, 240, 240)
                            }
                        },
                        'line': {
                            'visible': 0
                        }
                    }
                }
            }
        }
        # 标签2
        p_list = [s1_dict, ]

        # 跌涨柱颜色
        downupbars = [
            {'upbars': {
                'format': {
                    'fill': {
                        'ForeColor': {
                            'RGB': (0, 121, 68),
                        }
                    }
                }
            }
            },
            {'downbars': {
                'format': {
                    'fill': {
                        'ForeColor': {
                            'RGB': (157, 11, 11),
                        }
                    }
                }
            }
            }
        ]
        self.common_chart(sht_obj=sht_obj, chart_data_rng_obj=chart_rng, chart_name=chart_name,
                          chart_type=charttype, chart_size=chart_size,
                          chart_plotby=chart_plotby, gridline=False, chart_title=chart_title_dict,
                          legend=legend_dict, ticklabel=ticklabel_dict,
                          series=series_dict, point=p_list, updownbars=downupbars)

    # 日报 -- 组长业绩贡献率
    def reports_group_leader_rate(self, dq, wb_obj):
        start_date = self.start_date
        final_date = self.final_date
        df = ReportDataAsDf(dq, start_date, final_date).group_leader_rate()
        tab_color = 'R'

        sheet_name = '组长业绩贡献率'
        title = '{}各运营部{}'.format(dq, sheet_name)
        subtitles = self.subtitle(start_date, final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        # 图
        chart_name = sheet_name
        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(final_date, format='%m.%d')
        chart_title = '%s-%s%s推广小组组长业绩贡献率' % (title_date_start, title_date_final, dq)
        chart_plotby = c.xlColumns
        charttype = c.xlColumnStacked
        chart_size = (0, 170, 900, 400)
        chart_rng = sht_obj.Range('c3', 'g%d' % rs)
        # 标题
        chart_title_dict = self.chart_title(chart_title)
        # 图例
        legend_dict = self.chart_legend()
        # 坐标轴
        ticklabel_dict = self.chart_ticklabel()
        # 数据表
        datatable_dict = ticklabel_dict
        # 系列
        series_dict = [
            {-1: {
                'charttype': c.xlLine,
                'axisgroup': 2
            }
            }
        ]

        # 标签1
        s_dict = {
            '[-1,:]': {
                'datalabels': {
                    'font': {
                        'name': self.font_name,
                        'size': self.medium_size,
                        'bold': True,
                        'color': (91, 155, 213)
                    },
                    'position': 0,
                    'format': {
                        'fill': {
                            'forecolor': {
                                'rgb': (240, 240, 240)
                            }
                        },
                        'line': {
                            'visible': 0
                        }
                    }
                }
            }
        }
        # 标签2
        p_list = [s_dict, ]

        self.common_chart(sht_obj=sht_obj, chart_data_rng_obj=chart_rng, chart_name=chart_name,
                          chart_type=charttype, chart_size=chart_size,
                          chart_plotby=chart_plotby, gridline=False, chart_title=chart_title_dict,
                          legend=legend_dict, ticklabel=ticklabel_dict, datatable=datatable_dict,
                          series=series_dict, point=p_list)

    # 日报 -- 运营部白天夜间业绩对比
    def reports_team_day_evening(self, dq, wb_obj):
        start_date = self.start_date
        final_date = self.final_date
        df = ReportDataAsDf(dq, start_date, final_date).team_day_evening()
        tab_color = 'R'

        sheet_name = '运营部白天夜间业绩对比'
        title = sheet_name
        subtitles = self.subtitle(start_date, final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        # 图
        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(final_date, format='%m.%d')
        chart_title = '%s-%s %s地区各运营部业绩对比' % (title_date_start, title_date_final, dq)

        chart_name = sheet_name
        charttype = c.xlColumnStacked  # 类型
        chart_plotby = c.xlColumns
        chart_size = self.size_chart_datatable
        chart_data_rng = sht_obj.Range('c3', 'i%d' % rs)
        # 标题
        chart_title_dict = self.chart_title(chart_title)
        # 图例
        legend_dict = self.chart_legend()
        # 坐标轴
        ticklabel_dict = self.chart_ticklabel()
        # 数据表
        datatable_dict = self.chart_datatable()
        # 数据系列标签
        series_dict = [
            {-1: {
                'charttype': c.xlLine,
                'axisgroup': 2,
                'Format': {'Line': {'Weight': 2}},
            }
            },
            {(-2, -3): {
                'charttype': c.xlLine,
                'Format': {'Line': {'Weight': 2}},
            }
            },
            {2: {
                'charttype': c.xlLine,
                'Format': {'Line': {'Visible': 0}},
            }
            },
        ]
        # point
        s0_dict = {
            '[0,:]': {
                'datalabels': {
                    'font': {
                        'name': self.font_name,
                        'size': self.small_size,
                        'bold': True,
                        'color': (80, 137, 188)
                    },
                    'position': c.xlLabelPositionInsideBase,
                    'format': {
                        'line': {
                            'forecolor': {'rgb': (91, 155, 213), },
                            'visible': 0,
                            'weight': 1.1
                        },
                        'Fill': {
                            'ForeColor': {'RGB': (242, 242, 242)},
                        }
                    }
                }
            }
        }
        s4_dict = deepcopy(s0_dict)
        s4_dict['[4,-1]'] = s4_dict.pop('[0,:]')
        s4_dict['[4,-1]']['datalabels']['position'] = 0
        s4_dict['[4,-1]']['datalabels']['font']['color'] = (255, 0, 0)

        s5_dict = deepcopy(s4_dict)
        s5_dict['[5,:]'] = s5_dict.pop('[4,-1]')
        s5_dict['[5,:]']['datalabels']['font']['color'] = (98, 153, 62)

        p_dict = [s0_dict, s4_dict, s5_dict]

        self.common_chart(sht_obj=sht_obj, chart_data_rng_obj=chart_data_rng, chart_name=chart_name,
                          chart_type=charttype, chart_size=chart_size,
                          chart_plotby=chart_plotby, gridline=False, chart_title=chart_title_dict,
                          legend=legend_dict, datatable=datatable_dict,
                          ticklabel=ticklabel_dict, series=series_dict, point=p_dict)

    # 日报 -- 近6个月同期对比
    def reports_six_month(self, dq, wb_obj, df, tab_color):
        start_date = self.start_str_date_1
        final_date = self.final_date
        cols = list(df.columns)
        sheet_name = f'{cols[-7]}近6个月同期对比'
        title = sheet_name
        subtitles = self.subtitle(start_date, final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        css = cs - 8
        # 图1
        for i in range(2):
            n = (rs - 3) / 3
            cel = sht_obj.Range('a3')  # 类型
            cel_region, cel_colege, cel_team, cel_group = [cel.GetOffset(0, _) for _ in range(1, 5)]
            cel_start = sht_obj.Cells(3, cs - 5)
            cel_end = sht_obj.Cells(3, cs)

            chartrng_series = sht_obj.Range(cel_start, cel_end)
            chartrng_all = sht_obj.Range(cel_start.GetOffset(1 + n, 0), cel_end.GetOffset(2 * n, 0))
            chartrng_avg = chartrng_all.GetOffset(n, 0)
            chartrng_add = chartrng_series.GetOffset(3 * n + 1, 0)

            regionrng_series = cel
            regionrng_all = sht_obj.Range(regionrng_series.GetOffset(1 + n, 0), regionrng_series.GetOffset(2 * n, 0))
            regionrng_avg = regionrng_all.GetOffset(n, 0)
            regionrng_add = regionrng_series.GetOffset(3 * n + 1, 0)

            colegerng_series = cel_colege
            colegerng_all = sht_obj.Range(colegerng_series.GetOffset(1 + n, 0), colegerng_series.GetOffset(2 * n, 0))
            colegerng_avg = colegerng_all.GetOffset(n, 0)
            colegerng_add = colegerng_series.GetOffset(3 * n + 1, 0)

            teamrng_series = cel_team
            teamrng_all = sht_obj.Range(teamrng_series.GetOffset(1 + n, 0), teamrng_series.GetOffset(2 * n, 0))
            teamrng_avg = teamrng_all.GetOffset(n, 0)
            teamrng_add = teamrng_series.GetOffset(3 * n + 1, 0)

            grouprng_series = sht_obj.Range(cel_team, cel_group)
            grouprng_all = sht_obj.Range(grouprng_series.GetOffset(1 + n, 0), grouprng_series.GetOffset(2 * n, 0))
            grouprng_avg = grouprng_all.GetOffset(n, 0)
            grouprng_add = grouprng_series.GetOffset(3 * n + 1, 0)

            rng_all = [self.excel.Union(regionrng_series, chartrng_series, regionrng_all, chartrng_all),
                       self.excel.Union(colegerng_series, chartrng_series, colegerng_all, chartrng_all),
                       self.excel.Union(teamrng_series, chartrng_series, teamrng_all, chartrng_all),
                       self.excel.Union(grouprng_series, chartrng_series, grouprng_all, chartrng_all),
                       ]

            rng_avg = [self.excel.Union(regionrng_series, chartrng_series, regionrng_avg, chartrng_avg),
                       self.excel.Union(colegerng_series, chartrng_series, colegerng_avg, chartrng_avg),
                       self.excel.Union(teamrng_series, chartrng_series, teamrng_avg, chartrng_avg),
                       self.excel.Union(grouprng_series, chartrng_series, grouprng_avg, chartrng_avg),
                       ]

            data_rng = [rng_all, rng_avg]

            if n == 1 and i == 0 and css == 0:
                sht_obj.Cells(rs + 1, 1).Value = '增长量'
                for _ in range(2, cs - 5):
                    sht_obj.Cells(rs + 1, _).Value = sht_obj.Cells(rs, _).Value
                sht_obj.Cells(rs + 1, cs - 5).Value = 0
                for _ in range(cs - 4, cs + 1):
                    sht_obj.Cells(rs + 1, _).Value = sht_obj.Cells(5, _).Value - sht_obj.Cells(5, _ - 1).Value

                rng_add = [self.excel.Union(regionrng_add, chartrng_add),
                           self.excel.Union(colegerng_add, chartrng_add),
                           self.excel.Union(teamrng_add, chartrng_add),
                           self.excel.Union(grouprng_add, chartrng_add),
                           ]
                # 序列
                series_dict = [
                    {-1: {
                        'charttype': c.xlLine,
                        'Format': {'Line': {'Weight': 2}},
                    }
                    }, ]
                # point2
                s2_dict = {
                    '[1,:]': {
                        'datalabels': {
                            'font': {
                                'name': self.font_name,
                                'size': self.medium_size,
                                'bold': True,
                            },
                            'position': self.datalabel_position[0],
                        }
                    }
                }
                chart_data_rng = self.excel.Union(data_rng[i][css], rng_add[css])
            else:
                chart_data_rng = data_rng[i][css]
                series_dict = None
                s2_dict = None

            # title
            title_all = [f'QX{dq}-近6个月同期创量对比',
                         f'QX{dq}-近6个月{df.iloc[0, 2]}同期创量对比',
                         f'QX{dq}-近6个月各{cols[-7]}同期创量对比',
                         f'QX{dq}-近6个月各{cols[-7]}同期创量对比']
            title_avg = [f'QX{dq}-近6个月同期日人均创量对比',
                         f'QX{dq}-近6个月{df.iloc[0, 2]}同期日人均创量对比',
                         f'QX{dq}-近6个月各{cols[-7]}同期日人均创量对比',
                         f'QX{dq}-近6个月各{cols[-7]}同期日人均创量对比']

            chart_title_l = [title_all, title_avg]
            chart_title = chart_title_l[i][css]
            chart_name = chart_title
            charttype = self.chart_type[0]
            chart_plotby = self.chart_plotby[0]
            chart_style_num = self.chart_style_num[0]
            if css == 3:
                chart_size = self.size_chart_sixmonth_group[i]
            else:
                chart_size = self.size_chart_sixmonth[i]
            # 标题
            chart_title_dict = self.chart_title(chart_title)
            # 图例
            legend_dict = self.chart_legend()
            # 坐标轴
            ticklabel_dict = self.chart_ticklabel()
            # 数据表
            datatable_dict = self.chart_datatable()
            # point

            if s2_dict:
                s1_dict = {
                    '[0:-1,:]': {
                        'datalabels': {
                            'font': {
                                'name': self.font_name,
                                'size': self.medium_size,
                                'bold': True,
                            },
                            'position': self.datalabel_position[2],
                        }
                    }
                }
                p_dict = [s1_dict, s2_dict]
            else:
                s1_dict = {
                    '[:,:]': {
                        'datalabels': {
                            'font': {
                                'name': self.font_name,
                                'size': self.medium_size,
                                'bold': True,
                            },
                            'position': self.datalabel_position[2],
                        }
                    }
                }
                if css == 3:
                    s1_dict = {
                        '[:,:]': {
                            'datalabels': {
                                'font': {
                                    'name': self.font_name,
                                    'size': self.medium_size,
                                    'bold': True,
                                },
                                'position': self.datalabel_position[3],
                                'format': {'textframe2': {'orientation': 5}},
                            }
                        }
                    }
                p_dict = [s1_dict, ]

            self.common_chart(sht_obj=sht_obj, chart_data_rng_obj=chart_data_rng, chart_name=chart_name,
                              chart_type=charttype, chart_size=chart_size, chart_style_num=chart_style_num,
                              chart_plotby=chart_plotby, gridline=False, chart_title=chart_title_dict,
                              legend=legend_dict, datatable=datatable_dict,
                              ticklabel=ticklabel_dict, series=series_dict, point=p_dict)

    # 日报 -- 运营部晚间业绩对比
    def reports_evening(self, dq, wb_obj):
        start_date = self.start_date
        final_date = self.final_date
        df = ReportDataAsDf(dq, start_date, final_date).team_evening()
        tab_color = 'R'

        sheet_name = '各运营部夜间业绩对比'
        title = sheet_name
        subtitles = self.subtitle(start_date, final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)

        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(final_date, format='%m.%d')
        chart_title = '%s-%s %s' % (title_date_start, title_date_final, sheet_name)

        # 图
        chart_name = sheet_name
        charttype = c.xlColumnClustered
        chart_plotby = c.xlColumns
        chart_data_rng = sht_obj.Range(sht_obj.Range('c3'), sht_obj.Cells(rs, cs))
        # 标题
        chart_title_dict = self.chart_title(chart_title)
        # 图例
        legend_dict = self.chart_legend()
        # 坐标轴
        ticklabel_dict = self.chart_ticklabel()
        # 数据系列标签
        series_dict = [
            {-1: {'charttype': c.xlLine}}
        ]
        # point
        s1_dict = {
            '[0,:]': {
                'datalabels': {
                    'font': {
                        'name': self.font_name,
                        'size': self.medium_size
                    },
                }
            }
        }

        s2_dict = {
            '[-1,-1]': {
                'datalabels': {
                    'font': {
                        'name': self.font_name,
                        'size': 15,
                        'bold': True,
                        'color': 7434613
                    },
                    'position': 0,
                    'format': {'line': {
                        'ForeColor': {'RGB': (91, 155, 213)},
                        'visible': -1,
                        'weight': 1.1
                    }
                    }
                }
            }
        }
        p_dict = [s1_dict, s2_dict, ]
        self.common_chart(sht_obj=sht_obj, chart_data_rng_obj=chart_data_rng, chart_name=chart_name,
                          chart_type=charttype, chart_size=self.size_chart,
                          chart_plotby=chart_plotby, gridline=False, chart_title=chart_title_dict,
                          legend=legend_dict,
                          ticklabel=ticklabel_dict, series=series_dict, point=p_dict)

    # 日报 -- 运营部人员信息
    def reports_group_peoplelist(self, dq, wb_obj):
        start_date = self.start_date
        final_date = self.final_date
        df = ReportDataAsDf(dq, start_date, final_date).team_peoplelist()
        tab_color = 'G'
        sheet_name = '人员信息统计'
        title = sheet_name
        subtitles = self.subtitle(start_date, final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        rs1_rng = sht_obj.Cells(rs+1,1)
        rs1_d = {
            'value':'注：本月新人离职率=本月新进员工离职数/本月入职数     本月离职率=本月离职人数/（本月在职人数+本月离职人数）',
            'HorizontalAlignment':c.xlLeft,
            'VerticalAlignment':c.xlCenter,
             'font':{
                 'name':self.font_name,
                 'size':11,
                 'color':-16776961
            }
         }
        self.range_style(rs1_rng,**rs1_d)
        row3= sht_obj.Rows(3)
        row3_d = {
            'WrapText':True,
            'RowHeight':55
        }
        self.row_style(row3,**row3_d)
        sht_obj.Columns("D:D").ColumnWidth = 11
        sht_obj.Columns("F:Q").ColumnWidth = 8
        sht_obj.Columns("R:R").ColumnWidth = 12
        sht_obj.Columns("S:T").ColumnWidth = 7
        sht_obj.Columns("U:V").ColumnWidth = 8.5
        sht_obj.Columns("W:X").ColumnWidth = 13

        #合并单元格
        merge_rng = sht_obj.Range('a4',f'c{rs}')
        self.merge(sht_obj=sht_obj,rng=merge_rng,column_list=[1,2,3])
        #加粗行
        bold_rng= sht_obj.UsedRange
        self.bold(bold_rng,[4],tag='总计')
        sht_obj.Rows(f'4:{rs}').RowHeight = 25
        sht_obj.UsedRange.Select()

    # 日报 -- 目标完成率
    def reports_complete_rate(self, dq, wb_obj):
        start_date = self.start_date
        final_date = self.final_date
        df = ReportDataAsDf(dq, start_date, final_date).complete_rate_df()
        tab_color = 'G'

        sheet_name = '业绩目标完成率'
        title = f'QX{dq}-{sheet_name}'
        subtitles = self.subtitle(start_date, final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        # 合并单元格
        merge_rng = sht_obj.Range('a4', f'c{rs}')
        self.merge(sht_obj=sht_obj, rng=merge_rng, column_list=[1, 2, 3])
        # 加粗行
        bold_rng = sht_obj.UsedRange
        self.bold(bold_rng, [4], tag='总计')

        # 三角
        xltrigle_rng = sht_obj.Range('g4',f'g{rs}')
        self.xl3Triangles(xltrigle_rng,columnlist=[1])
        # 数据条
        databar_rng = xltrigle_rng.GetOffset(0,1)
        self.xlConditionValueNumber(databar_rng,columnlist=[1])

        sht_obj.UsedRange.Cells.Columns.ColumnWidth = 17
        sht_obj.Rows(f'4:{rs}').RowHeight = 25

        sht_obj.UsedRange.Select()

    # 日报 -- 组长组员日人均业绩
    #todo
    def reports_group_leader_menber(self, dq, wb_obj):
        start_date = self.start_date
        final_date = self.final_date
        df = ReportDataAsDf(dq, start_date, final_date).complete_rate_df()
        tab_color = 'G'

        sheet_name = '业绩目标完成率'
        title = f'QX{dq}-{sheet_name}'
        subtitles = self.subtitle(start_date, final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        # 合并单元格
        merge_rng = sht_obj.Range('a4', f'c{rs}')
        self.merge(sht_obj=sht_obj, rng=merge_rng, column_list=[1, 2, 3])
        # 加粗行
        bold_rng = sht_obj.UsedRange
        self.bold(bold_rng, [4], tag='总计')

        # 三角
        xltrigle_rng = sht_obj.Range('g4',f'g{rs}')
        self.xl3Triangles(xltrigle_rng,columnlist=[1])
        # 数据条
        databar_rng = xltrigle_rng.GetOffset(0,1)
        self.xlConditionValueNumber(databar_rng,columnlist=[1])

        sht_obj.UsedRange.Cells.Columns.ColumnWidth = 17
        sht_obj.Rows(f'4:{rs}').RowHeight = 25

        sht_obj.UsedRange.Select()
if __name__ == '__main__':
    # Reports('2020/5/22').report_component()
    Reports('2020/5/27').reports_day()
