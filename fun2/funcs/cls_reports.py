#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/21 9:41:28
# @Author  : HouWk
# @Site    : 
# @File    : cls_reports.py
# @Software: PyCharm
from cls_excel import Excel
from copy import deepcopy
from win32com.client import constants as c  # 旨在直接使用VBA常数
from cls_sqlserver import ReportDataAsDf, Component
from fun_date import get_str_date
from fun_os import create_folder_date


class Reports(Excel):
    def __init__(self, final_date):
        super().__init__()

        # 报表储存路径
        self.root_path = r'e:\python_Reports'

        # 终止日期
        self.final_date = final_date

        # 数据表tab颜色
        self.tab_colors = ['R', 'DR', 'G', 'DG','Y']

        # 分量地区
        self.component_dqs = ['燕郊'] #['济南', '燕郊', '成都']
        # 分量图线条颜色
        self.RGBs = [(255, 255, 0), (0, 112, 192), (0, 176, 80), (255, 0, 0), (255, 255, 255)]  # 黄 蓝 绿 红 白
        self.component_RGB = self.RGBs[:4]

        # 其余报表地区
        # self.common_dqs = ['保定', '济南']
        self.common_dqs = ['保定']

        # 图大小
        self.size_long_chart = (0, 170, 1200, 500)
        self.size_chart = (0, 170, 900, 400)
        self.size_chart_datatable = (0, 170, 900, 500)

        self.size_chart_sixmonth = [(600, 0, 900, 400), (600, 401, 900, 400)]
        self.size_chart_sixmonth_group = [(600, 0, 900, 600), (600, 601, 900, 600)]

        # charttype 0-xlColumnClustered柱状图 1-折线图 2-折线跌涨柱  3-柱状堆积图
        self.chart_type = [c.xlColumnClustered, c.xlLine, c.xlLineMarkers, c.xlColumnStacked]

        # chartplotby 0-xlRows  1-xlColumns
        self.chart_plotby = [c.xlRows, c.xlColumns]

        # chart_style_num  0-233
        self.chart_style_num = [233, ]

        # datalabelposition [0、1] 折线  0-数据点上方 1-数据点下方  [2、3] 柱形图 2-上边缘上方 3-上边缘下方
        self.datalabel_position = [c.xlLabelPositionAbove, c.xlLabelPositionBelow,
                                   c.xlLabelPositionOutsideEnd, c.xlLabelPositionInsideEnd]

    def reports_morning_evening(self,hours):
        self.screen_updating(False)  # 关闭屏幕刷新
        wb_obj = self.excel.Workbooks.Add()  # 创建excel wb
        for dq in self.common_dqs:
            df, start_date, final_date = Component(dq=dq, final_date=self.final_date).component_df()  # 数据帧
            # 表
            sheet_name = '%s近%d月日人均分量' % (dq, 4)
            # 区域标题
            r_title_v = sheet_name
            # 区域副标题
            r_subtitles_v = self.subtitle(start_date, final_date)
            # 区域赋值 并 设定格式
            sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=self.tab_colors[4],
                                                df=df,
                                                title_v=r_title_v, subtitles_v=r_subtitles_v)
            # 图
            chart_size = self.size_long_chart
            chart_type = self.chart_type[1]
            chart_name = sheet_name
            chart_plotby = self.chart_plotby[0]
            chart_style_num = self.chart_style_num[0]
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
        path = create_folder_date(self.root_path, self.final_date)  # 创建目标文件夹
        str_date = get_str_date(self.final_date, '%Y%m%d')
        wb_name = '%s日报_人均分量' % (str_date)
        self.workbook_save(wookbook_obj=wb_obj, name=wb_name, path=path)
        self.screen_updating(True)  # 开启屏幕刷新


    # 报表 分量
    def report_component(self):
        self.screen_updating(False)  # 关闭屏幕刷新
        wb_obj = self.excel.Workbooks.Add()  # 创建excel wb
        for dq in self.component_dqs:
            df, start_date, final_date = Component(dq=dq, final_date=self.final_date).component_df()  # 数据帧
            # 表
            sheet_name = '%s近%d月日人均分量' % (dq, 4)
            # 区域标题
            r_title_v = sheet_name
            # 区域副标题
            r_subtitles_v = self.subtitle(start_date, final_date)
            # 区域赋值 并 设定格式
            sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=self.tab_colors[4], df=df,
                                                title_v=r_title_v, subtitles_v=r_subtitles_v)
            # 图
            chart_size = self.size_long_chart
            chart_type = self.chart_type[1]
            chart_name = sheet_name
            chart_plotby = self.chart_plotby[0]
            chart_style_num = self.chart_style_num[0]
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
        path = create_folder_date(self.root_path, self.final_date)  # 创建目标文件夹
        str_date = get_str_date(self.final_date, '%Y%m%d')
        wb_name = '%s日报_人均分量' % (str_date)
        self.workbook_save(wookbook_obj=wb_obj, name=wb_name, path=path)
        self.screen_updating(True)  # 开启屏幕刷新

    # 日报 总
    def reports_day(self):
        for dq in self.common_dqs:
            self.dq = dq
            self.screen_updating(False)  # 关闭屏幕刷新

            # self.screen_updating(True)  # 开启屏幕刷新

            wb_obj = self.excel.Workbooks.Add()  # 创建excel wb

            # 数据帧类
            cls_df = ReportDataAsDf(dq, self.final_date)

            # T
            # 日报4--【(1-3)/4】--进群率和注册率
            group_er_en, team_er_en, colege_er_en, region_er_en, start_date_er_en = cls_df.erollment_and_group_entry_rate_df()
            self.report_erollment_and_group_entry_rate(wb_obj=wb_obj, tab_color=None, df=region_er_en,
                                                       start_date=start_date_er_en)
            self.report_erollment_and_group_entry_rate(wb_obj=wb_obj, tab_color=self.tab_colors[0], df=colege_er_en,
                                                       start_date=start_date_er_en)
            self.report_erollment_and_group_entry_rate(wb_obj=wb_obj, tab_color=self.tab_colors[1], df=team_er_en,
                                                       start_date=start_date_er_en)
            # 日报3-近4月趋势
            dfs_region, dfs_colege, dfs_team, start_date_4month = cls_df.trend_4monthes_df()
            self.report_4monthes_trend_region(wb_obj=wb_obj,dfs=dfs_region,start_date=start_date_4month)
            self.report_4monthes_trend_colege(wb_obj=wb_obj,dfs=dfs_colege,start_date=start_date_4month)
            self.report_4monthes_trend_team(wb_obj=wb_obj,dfs=dfs_team,start_date=start_date_4month)
            # 日报-上月同期对比
            self.reports_vs_last_month(wb_obj=wb_obj,cls_df=cls_df)
            # 日报-时间消耗率与完成率
            self.reports_complete_rate_time_rate(wb_obj=wb_obj,cls_df=cls_df)
            # 日报 -- 组长业绩贡献率
            self.reports_group_leader_rate(wb_obj=wb_obj,cls_df=cls_df)
            # 日报2--【1/2】 -- 运营部或小组 白天夜间业绩对比
            df_group_day_evening, df_team_day_evening, start_date_team_day_evening = cls_df.group_team_day_evening()
            self.reports_team_day_evening(wb_obj=wb_obj, df=df_team_day_evening, start_date=start_date_team_day_evening,
                                          r_type=0)
            # 日报4--【(1-3)/4】-近6个月同期对比
            sixmonth_group, sixmonth_team, sixmonth_colege, sixmonth_region,start_date_sixmonth = cls_df.team_six_month()
            self.reports_six_month(dq=dq, wb_obj=wb_obj, df=sixmonth_region, tab_color=self.tab_colors[4],start_date=start_date_sixmonth)
            self.reports_six_month(dq=dq, wb_obj=wb_obj, df=sixmonth_colege, tab_color=self.tab_colors[0],start_date=start_date_sixmonth)
            self.reports_six_month(dq=dq, wb_obj=wb_obj, df=sixmonth_team, tab_color=self.tab_colors[0],start_date=start_date_sixmonth)
            #日报2-【1/2】- 夜间业绩对比
            df_group_evening,df_team_evening, start_date_evening = cls_df.evening()
            self.reports_evening(wb_obj=wb_obj,df=df_team_evening,start_date=start_date_evening,tab_color=self.tab_colors[0],type=0)

            # G
            # 日报-人员信息统计
            self.reports_group_peoplelist(wb_obj=wb_obj,cls_df=cls_df)
            # 日报-业绩目标完成率
            self.reports_complete_rate(wb_obj=wb_obj,cls_df=cls_df)
            # 日报4--【4/4】--进群率和注册率
            self.report_erollment_and_group_entry_rate(wb_obj=wb_obj, tab_color=self.tab_colors[2], df=group_er_en,
                                                       start_date=start_date_er_en)
            # 日报-组长组员日人均业绩
            self.reports_group_leader_member(wb_obj=wb_obj, cls_df=cls_df)
            # 日报 -- 当日业绩统计
            self.reports_people_performance(wb_obj=wb_obj, cls_df=cls_df)
            # 日报 -- 推广专员月内日均创量排名
            self.reports_people_month_avg_rank(wb_obj=wb_obj,  cls_df=cls_df)
            # 日报2 -- 1-推广小组组长当月与上月日均业绩对比 2-推广组组长当月业绩排名
            self.reports_groupleader_vs_lastmonth_avg(wb_obj=wb_obj, cls_df=cls_df) # 两个报表
            #日报2-【2/2】- 夜间业绩对比
            self.reports_evening(wb_obj=wb_obj,df=df_group_evening,start_date=start_date_evening,tab_color=self.tab_colors[2],type=1)
            # 日报 -日阶段完成任务次数
            self.reports_days_finish(wb_obj=wb_obj, cls_df=cls_df)
            #日报 - 推广小组日均排名
            self.reports_groups_avg_rank_peoples(wb_obj=wb_obj, cls_df=cls_df)
            # 日报2--【2/2】 -- 小组白天夜间业绩对比
            self.reports_team_day_evening(wb_obj=wb_obj, df=df_group_day_evening, start_date=start_date_team_day_evening,
                                          r_type=1)
            # 日报4--【4/4】-近6个月同期对比
            self.reports_six_month(dq=dq, wb_obj=wb_obj, df=sixmonth_group, tab_color=self.tab_colors[2],start_date=start_date_sixmonth)

            wb_obj.Sheets(1).Delete()
            wb_obj.Sheets(1).Select()

            # 保存文件
            path = create_folder_date(self.root_path, self.final_date)  # 创建目标文件夹
            str_date = get_str_date(self.final_date, '%Y%m%d')
            wb_name = '%s%s日报' % (str_date, dq)
            self.workbook_save(wookbook_obj=wb_obj, name=wb_name, path=path)
            self.screen_updating(True)  # 开启屏幕刷新

    # 日报 - 推广进群率与注册率统计
    def report_erollment_and_group_entry_rate(self, wb_obj, tab_color, df, start_date):
        dep = df.columns[-6]
        sheet_name = '%s推广进群率与注册率统计' % dep
        title = sheet_name
        subtitles = self.subtitle(start_date, self.final_date)
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

    # 日报 - 近4个月趋势—地区
    def report_4monthes_trend_region(self, wb_obj, dfs, start_date):
        for df in dfs:
            dep = df.iloc[0, 0]
            sheet_name = '%s日创量趋势' % dep
            title = sheet_name
            subtitles = self.subtitle(start_date, self.final_date)
            sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=self.tab_colors[4], df=df,
                                                title_v=title, subtitles_v=subtitles)
            chart_data_rng = sht_obj.Range(sht_obj.Range('b3'), sht_obj.Cells(rs, cs))
            self.chart_4monthes_trend(dep, sht_obj, chart_data_rng)

    # 日报 - 近4个月趋势-学院
    def report_4monthes_trend_colege(self, wb_obj, dfs, start_date):
        for df in dfs:
            dep = ''.join(list(df.iloc[0, :2]))
            sheet_name = '%s日创量趋势' % dep
            title = sheet_name
            subtitles = self.subtitle(start_date, self.final_date)
            sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=self.tab_colors[0], df=df,
                                                title_v=title, subtitles_v=subtitles)
            chart_data_rng = sht_obj.Range(sht_obj.Range('c3'), sht_obj.Cells(rs, cs))
            self.chart_4monthes_trend(dep, sht_obj, chart_data_rng)

    # 日报 - 近4个月趋势-运营部
    def report_4monthes_trend_team(self, wb_obj, dfs, start_date):
        for df in dfs:
            dep = ''.join(list(df.iloc[0, :3]))
            sheet_name = '%s日创量趋势' % dep
            title = sheet_name
            subtitles = self.subtitle(start_date, self.final_date)
            sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=self.tab_colors[0], df=df,
                                                title_v=title, subtitles_v=subtitles)
            chart_data_rng = sht_obj.Range(sht_obj.Range('d3'), sht_obj.Cells(rs, cs))
            self.chart_4monthes_trend(dep, sht_obj, chart_data_rng)

    # 日报 -- 与上月同期业绩对比
    def reports_vs_last_month(self, wb_obj, cls_df):
        df, start_date = cls_df.vs_last_month_df()
        sheet_name = '与上月同期业绩对比'
        title = sheet_name
        subtitles = self.subtitle(start_date, self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=self.tab_colors[0], df=df,
                                            title_v=title, subtitles_v=subtitles)

        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(self.final_date, format='%m.%d')
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

        p_dict = [s1_s2_dict, s3_dict, s4_dict]

        self.common_chart(sht_obj=sht_obj, chart_data_rng_obj=chart_data_rng, chart_name=chart_name,
                          chart_type=charttype, chart_size=self.size_chart,
                          chart_plotby=chart_plotby, gridline=False, chart_title=chart_title_dict,
                          legend=legend_dict,
                          ticklabel=ticklabel_dict, series=series_dict, point=p_dict)

    # 日报 -- 完成率与时间消耗率
    def reports_complete_rate_time_rate(self, wb_obj, cls_df):
        df, start_date = cls_df.complete_rate_time_rate_df()
        sheet_name = '完成率与时间消耗率'
        title = '各运营部{}'.format(sheet_name)
        subtitles = self.subtitle(start_date, self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=self.tab_colors[0], df=df,
                                            title_v=title, subtitles_v=subtitles)
        # 图
        chart_name = sheet_name
        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(self.final_date, format='%m.%d')
        chart_title = '%s-%s完成率与时间消耗率涨跌图' % (title_date_start, title_date_final)
        chart_plotby = self.chart_plotby[1]  # c.xlColumns
        charttype = self.chart_type[2]  # c.xlLineMarkers
        chart_size = self.size_chart  # (0, 170, 900, 400)
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
    def reports_group_leader_rate(self, wb_obj, cls_df):
        df, start_date = cls_df.group_leader_rate()
        sheet_name = '组长业绩贡献率'
        title = sheet_name
        subtitles = self.subtitle(start_date, self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=self.tab_colors[0], df=df,
                                            title_v=title, subtitles_v=subtitles)
        # 图
        chart_name = sheet_name
        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(self.final_date, format='%m.%d')
        chart_title = '%s-%s 推广小组组长业绩贡献率' % (title_date_start, title_date_final)
        chart_plotby = self.chart_plotby[1]  # c.xlColumns
        charttype = self.chart_type[3]  # c.xlColumnStacked
        chart_size = self.size_chart  # (0, 170, 900, 400)
        chart_rng = sht_obj.Range('c3', 'g%d' % rs)
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
    def reports_team_day_evening(self, wb_obj, df, start_date, r_type):
        if r_type == 0:
            data_type = '运营部'
            tab_color = self.tab_colors[0]
        elif r_type == 1:
            data_type = '推广小组'
            tab_color = self.tab_colors[2]
        else:
            raise Exception('r_type Error')

        sheet_name = f'{data_type}白天夜间业绩对比'
        title = sheet_name
        subtitles = self.subtitle(start_date, self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        # 图
        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(self.final_date, format='%m.%d')
        chart_title = '%s-%s 各%s业绩对比' % (title_date_start, title_date_final,data_type)

        chart_name = sheet_name
        charttype = self.chart_type[3]  # c.xlColumnStacked  # 类型
        chart_plotby = self.chart_plotby[1]  # c.xlColumns
        chart_size = self.size_chart_datatable
        chart_data_rng = sht_obj.Range('c3', sht_obj.Cells(rs,cs-1))
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
    def reports_six_month(self, dq, wb_obj, df, tab_color, start_date):
        cols = list(df.columns)
        sheet_name = f'{cols[-7]}近6个月同期对比'
        title = sheet_name
        subtitles = self.subtitle(start_date, self.final_date)
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

            if n == 1 and i == 0 and (css == 0 or css == 1):
                cell = sht_obj.Cells(rs+5, 1)
                if css == 0:
                    cell.Value = '只发日人均创量图到全国群'
                elif css == 1:
                    cell.Value = '只发同期创量图到业务群'

                cell.Font.Bold = True
                cell.Font.Size = 20
                cell.Font.Color = 255

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
    def reports_evening(self, wb_obj, df, start_date, tab_color, type):
        types = ['运营部', '推广小组']

        sheet_name = f'各{types[type]}夜间业绩对比'
        title = sheet_name
        subtitles = self.subtitle(start_date, self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)

        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(self.final_date, format='%m.%d')
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
                        'size': self.medium_size,
                        'bold': True,
                        'color': (255, 0, 0)
                    },
                    'position': 0,
                    'format': {
                        'line': {
                            'ForeColor': {'RGB': (91, 155, 213)},
                            'visible': -1,
                            'weight': 1.1
                        },
                        'fill': {
                            'forecolor': {'rgb': (240, 240, 240)}
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
    def reports_group_peoplelist(self, wb_obj, cls_df):
        df, start_date = cls_df.team_peoplelist()
        tab_color = self.tab_colors[2]
        sheet_name = '人员信息统计'
        title = sheet_name
        subtitles = self.subtitle(start_date, self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        rs1_rng = sht_obj.Cells(rs + 1, 1)
        rs1_d = {
            'value': '注：本月新人离职率=本月新进员工离职数/本月入职数     本月离职率=本月离职人数/（本月在职人数+本月离职人数）',
            'HorizontalAlignment': c.xlLeft,
            'VerticalAlignment': c.xlCenter,
            'font': {
                'name': self.font_name,
                'size': 11,
                'color': -16776961
            }
        }
        self.range_style(rs1_rng, **rs1_d)
        row3 = sht_obj.Rows(3)
        row3_d = {
            'WrapText': True,
            'RowHeight': 55
        }
        self.row_style(row3, **row3_d)
        sht_obj.Columns("D:D").ColumnWidth = 11
        sht_obj.Columns("F:Q").ColumnWidth = 8
        sht_obj.Columns("R:R").ColumnWidth = 12
        sht_obj.Columns("S:T").ColumnWidth = 7
        sht_obj.Columns("U:V").ColumnWidth = 8.5
        sht_obj.Columns("W:X").ColumnWidth = 13

        # 合并单元格
        merge_rng = sht_obj.Range('a4', f'c{rs}')
        self.merge(sht_obj=sht_obj, rng=merge_rng, column_list=[1, 2, 3])
        # 加粗行
        bold_rng = sht_obj.UsedRange
        self.bold(bold_rng, [4], tag='总计')
        sht_obj.Rows(f'4:{rs}').RowHeight = 25
        sht_obj.UsedRange.Select()

    # 日报 -- 目标完成率
    def reports_complete_rate(self, wb_obj, cls_df):
        df, start_date = cls_df.complete_rate_df()
        tab_color = self.tab_colors[3]
        sheet_name = '业绩目标完成率'
        title = f'QX{self.dq}-{sheet_name}'
        subtitles = self.subtitle(start_date, self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        # 合并单元格
        merge_rng = sht_obj.Range('a4', f'c{rs}')
        self.merge(sht_obj=sht_obj, rng=merge_rng, column_list=[1, 2, 3])
        # 加粗行
        bold_rng = sht_obj.UsedRange
        self.bold(bold_rng, [4], tag='总计')

        # 三角
        xltrigle_rng = sht_obj.Range('g4', f'g{rs}')
        self.xl3Triangles(xltrigle_rng, columnlist=[1])
        # 数据条
        databar_rng = xltrigle_rng.GetOffset(0, 1)
        self.xlCondition_percent_databar(databar_rng, columnlist=[1])

        sht_obj.UsedRange.Cells.Columns.ColumnWidth = 17
        sht_obj.Rows(f'4:{rs}').RowHeight = 25

        sht_obj.UsedRange.Select()

    # 日报 -- 组长组员日人均业绩
    def reports_group_leader_member(self, wb_obj, cls_df):
        df, start_date = cls_df.group_leader_member()
        tab_color = self.tab_colors[2]
        sheet_name = '组长和组员日人均业绩合格率对比'
        title = sheet_name
        subtitles = self.subtitle(start_date, self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles)
        # 合并单元格
        merge_rng = sht_obj.Range('a4', f'c{rs}')
        self.merge(sht_obj=sht_obj, rng=merge_rng, column_list=[1, 2, 3])
        # 图
        title_date_start = get_str_date(start_date, format='%m.%d')
        title_date_final = get_str_date(self.final_date, format='%m.%d')
        chart_title = '%s-%s推广小组组长和组内人员日人均合格率对比图' % (title_date_start, title_date_final)

        # 图
        chart_name = sheet_name
        charttype = c.xlColumnClustered  # 类型
        chart_plotby = c.xlColumns
        chart_data_rng = sht_obj.Range(sht_obj.Range('c3'), sht_obj.Cells(rs, 8))
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

        p_dict = [s1_s2_dict, s3_dict, s4_dict]

        self.common_chart(sht_obj=sht_obj, chart_data_rng_obj=chart_data_rng, chart_name=chart_name,
                          chart_type=charttype, chart_size=self.size_chart,
                          chart_plotby=chart_plotby, gridline=False, chart_title=chart_title_dict,
                          legend=legend_dict,
                          ticklabel=ticklabel_dict, series=series_dict, point=p_dict)

    # 日报 -- 当日业绩统计
    def reports_people_performance(self, wb_obj, cls_df):
        df, start_date = cls_df.people_performance()
        sheet_name = '推广专员当日业绩统计'
        title = sheet_name
        subtitles = self.subtitle(self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=self.tab_colors[2], df=df,
                                            title_v=title, subtitles_v=subtitles, shadow=True)
        # 合并单元格
        interior_rng = sht_obj.Range('i4', f'i{rs}')
        for rng in interior_rng:
            if rng.Value < 30:
                rng.Interior.Color = 11854022  # 蓝色
            elif rng.Value < 50:
                rng.Interior.Color = 8781823  # 绿色
            elif rng.Value < 100:
                rng.Interior.Color = 7697919  # 红色
            else:
                rng.Interior.Color = 49407  # 橙色

    # 日报 -- 推广专员月内日均创量排名
    def reports_people_month_avg_rank(self, wb_obj, cls_df):
        df, start_date = cls_df.people_month_avg_rank()
        tab_color = self.tab_colors[2]

        sheet_name = '推广专员月内日均业绩排名'
        title = sheet_name
        subtitles = self.subtitle(self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles, shadow=True)

    # 日报2 -- 日报2 -- 1-推广小组组长当月与上月日均业绩对比 2-推广组组长当月业绩排名
    def reports_groupleader_vs_lastmonth_avg(self, wb_obj, cls_df):
        df1, df2, avg_line, start_date = cls_df.groupleader_vs_lastmonth_avg()
        tab_color = self.tab_colors[2]

        sheet_name1 = '推广小组组长当月与上月日均业绩对比'
        title1 = sheet_name1
        subtitles1 = self.subtitle(self.final_date)
        sht_obj1, rs1, cs1 = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name1, tab_color=tab_color, df=df1,
                                               title_v=title1, subtitles_v=subtitles1, shadow=True)

        sheet_name2 = '推广组组长当月业绩排名'
        title2 = f'{sheet_name2}(均量平均线={avg_line})'
        subtitles2 = self.subtitle(start_date, self.final_date)
        sht_obj2, rs2, cs2 = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name2, tab_color=tab_color, df=df2,
                                               title_v=title2, subtitles_v=subtitles2, shadow=True)

    #日报 -日阶段完成任务次数
    def reports_days_finish(self,wb_obj,cls_df):
        df, start_date = cls_df.days_finish()
        tab_color = self.tab_colors[2]
        sheet_name = '推广小组日阶段完成任务次数统计'
        title = sheet_name
        subtitles = self.subtitle(start_date,self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                               title_v=title, subtitles_v=subtitles, shadow=True)

        for cel in sht_obj.Range('e4',sht_obj.Cells(rs,cs-4)):
            if cel.Value > 750:
                cel.Interior.Color = 49407 #橙色
                if cel.Value > 1000:
                    cel.Interior.Color = 8781823 #黄色

        percent_databar_rng = sht_obj.Range(sht_obj.Cells(4,cs-1),sht_obj.Cells(rs,cs-1))
        self.xlCondition_percent_databar(percent_databar_rng, columnlist=[1])

        maxmin_databar_rng1 = percent_databar_rng.GetOffset(0,-2) #750
        maxmin_databar_rng2 = percent_databar_rng.GetOffset(0,-1) #1000
        self.xlCondition_maxmin_databar(rng=maxmin_databar_rng1,columnlist=[1],color=8700771)
        self.xlCondition_maxmin_databar(rng=maxmin_databar_rng2,columnlist=[1],color=2668287)

    # 日报 - 推广小组日均业绩排名详情
    def reports_groups_avg_rank_peoples(self, wb_obj,cls_df):
        df, start_date = cls_df.groups_avg_rank_peoples()
        tab_color = self.tab_colors[2]

        sheet_name = '推广小组日均业绩排名'
        title = sheet_name
        subtitles = self.subtitle(start_date,self.final_date)
        sht_obj, rs, cs = self.common_sheet(wb_obj=wb_obj, sht_name=sheet_name, tab_color=tab_color, df=df,
                                            title_v=title, subtitles_v=subtitles, shadow=True)

if __name__ == '__main__':
    r = Reports('2020/6/7')
    # r.report_component()
    r.reports_day()
    # Reports('2020/6/2').report_component()
    # Reports('2020/6/7').reports_day()
