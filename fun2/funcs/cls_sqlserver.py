#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/20 9:05:01
# @Author  : HouWk
# @Site    : 
# @File    : cls_sqlserver.py
# @Software: PyCharm
import datetime
from copy import deepcopy
import numpy as np
import pyodbc
from pandas import DataFrame
import pandas as pd
from cls_date import MyDate, get_str_date
from fun_df import format_series_percentage, reindex_cols, fun_groupby, cartesian_product_df, \
    fill_na, FormatSeriesOfDf, merge_many_dfs


class Odbc:
    def __init__(self, dr, sv, db, un=None, pw=None):
        self.dr = dr  # driver
        self.sv = sv  # 服务器名称
        self.db = db  # '数据库名称
        self.un = un  # '数据库连接用户名
        self.pw = pw  # '数据库连接密码

    def conn_str(self):
        if self.un and self.pw:
            conn_str = r'DRIVER={};SERVER={};DATABASE={};UID={};PWD={}'.format(
                self.dr, self.sv, self.db, self.un, self.pw)
        else:
            conn_str = r'DRIVER={};SERVER={};DATABASE={};Trusted_Connection=Yes'.format(
                self.dr, self.sv, self.db)
        return conn_str

    def operate_DB(self, sql):  # 数据库
        con_str = self.conn_str()
        conn = pyodbc.connect(con_str)
        try:
            sql_data = pd.read_sql(sql, conn)
            return DataFrame(sql_data)
        except:
            cursor = conn.cursor()
            cursor.execute(sql)
            conn.commit()


class MySqlServer(Odbc):
    def __init__(self, dq, start_date=None, final_date=None, table_name='Tg'):
        dr = 'SQL Server Native Client 11.0'  # driver
        # sv = "192.168.1.43"  # 服务器名称
        sv = 'localhost'  # 服务器名称
        db = "QXT"  # '数据库名称
        # un = "sa"  # '数据库连接用户名
        # pw = "houweikang123"  # '数据库连接密码
        super().__init__(dr=dr, sv=sv, db=db)

        self.dq = dq
        self.start_date = start_date
        self.final_date = final_date
        self.table_name = table_name

        self.dict_substituet = {
            '1战队': '运营一部',
            '2战队': '运营二部',
            '3战队': '运营三部',
            '4战队': '运营四部',
            '5战队': '运营五部',
            '6战队': '运营六部',
            '7战队': '运营七部',
            '8战队': '运营八部',
        }

    def get_Peoplelist(self):
        sql = '''SELECT [员工姓名]
                  ,[员工工号] [工号]
                  ,[员工岗位]
                  ,[所属学院] [学院]
                  ,[所属部门] 
                  ,[展翅账号]
                  ,[主站账号]
                  ,[接量类型]
                  ,[入职时间]
                  ,[账号状态]
                  ,[状态]
                  ,[地区]
                  ,[战队] [运营部]
                  ,[小组]
                  ,[离职日期]
              FROM [QXT].[dbo].[PeopleList]
              WHERE [地区] = '%s' ''' % self.dq
        df = self.operate_DB(sql)
        df['运营部'] = df['运营部'].map(self.dict_substituet)
        return df

    def get_tg_data(self):
        '''
        新增了 [提交日期] type(date) ; [提交小时] type(int)
        :param table_name: Tg or HourTg
        :return:
        '''
        sql = '''SELECT [账号]
                  ,[账号类型]
                  ,[班级代号]
                  ,[群号/微信号]
                  ,[推广专员]
                  ,[推广专员-所属学院] [学院]
                  ,[推广专员-所属地区] [地区]
                  ,[推广专员-所属战队] [运营部]
                  ,[推广专员-所属小组] [小组]
                  ,[推广专员-所属岗位] [岗位]
                  ,[推广专员-员工号] [工号]
                  ,[课程顾问]
                  ,[课程顾问-所属学院]
                  ,[课程顾问-所属地区]
                  ,[课程顾问-所属战队]
                  ,[课程顾问-所属小组]
                  ,[课程顾问-所属岗位]
                  ,[课程顾问-员工号]
                  ,[提交时间]
                  ,[进群状态]
                  ,[进群时间]
                  ,[申请进群时间]
                  ,[是否注册]
                  ,[开通课程]
                  ,[量来源]
                  ,CAST([提交时间] as date) [提交日期]
                  ,DATEPART(hh, [提交时间]) [提交小时]
              FROM [QXT].[dbo].[{}]
              WHERE [推广专员-所属地区] = '{}'
                    and CAST([提交时间] as date) between '{}' and '{}' '''.format(self.table_name, self.dq, self.start_date,
                                                                              self.final_date)
        df = self.operate_DB(sql)
        df['运营部'] = df['运营部'].map(self.dict_substituet)
        return df

    def get_tg_data_where_guwen(self):
        '''
        新增了 [提交日期] type(date) ; [提交小时] type(int)
        :param table_name: Tg or HourTg
        :return:
        '''
        sql = '''SELECT [账号]
                    ,[账号类型]
                    ,[班级代号]
                    ,[群号/微信号]
                    ,[推广专员]
                    ,[推广专员-所属学院] [学院]
                    ,[推广专员-所属地区] [地区]
                    ,[推广专员-所属战队] [运营部]
                    ,[推广专员-所属小组] [小组]
                    ,[推广专员-所属岗位] [岗位]
                    ,[推广专员-员工号] [工号]
                    ,[课程顾问]
                    ,[课程顾问-所属学院]
                    ,[课程顾问-所属地区]
                    ,[课程顾问-所属战队]
                    ,[课程顾问-所属小组]
                    ,[课程顾问-所属岗位]
                    ,[课程顾问-员工号]
                    ,[提交时间]
                    ,[进群状态]
                    ,[进群时间]
                    ,[申请进群时间]
                    ,[是否注册]
                    ,[开通课程]
                    ,[量来源]
                    ,CAST([提交时间] as date) [提交日期]
                    ,DATEPART(hh, [提交时间]) [提交小时]
                FROM [QXT].[dbo].[{}]
                WHERE [课程顾问-所属地区] = '{}'
                      and CAST([提交时间] as date) between '{}' and '{}' '''.format(self.table_name, self.dq,
                                                                                self.start_date,
                                                                                self.final_date)
        df = self.operate_DB(sql)
        df['运营部'] = df['运营部'].map(self.dict_substituet)
        return df

    def get_tg_groups_groupby_date(self):

        '''
        :return:'地区', '学院', '运营部', '小组', '日期', '创量'
        '''
        cols = ['地区', '学院', '运营部', '小组', '提交日期']
        df = self.get_tg_data()[cols]
        cols = ['地区', '学院', '运营部', '小组', '日期']
        df.rename(columns={'提交日期': '日期'}, inplace=True)
        cond = df['小组'].str.endswith('组')
        df = df[cond]
        df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '创量'})
        return df

    def get_tg_groups_groupby_date_hours(self, start_hour, end_hour):
        '''
        :return:'地区', '学院', '运营部', '小组', '日期','提交小时', '业绩'
        '''
        cols = ['地区', '学院', '运营部', '小组', '提交日期', '提交小时']
        df = self.get_tg_data()[cols]
        cond = df['小组'].str.endswith('组')
        cond1 = (end_hour >= df['提交小时'])
        cond2 = (df['提交小时'] >= start_hour)
        df = df[cond & cond1 & cond2]
        df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '业绩'})
        return df

    def get_tg_groups_groupby(self):
        '''
        :return:'地区', '学院', '运营部', '小组', '创量'
        '''
        cols = ['地区', '学院', '运营部', '小组']
        df = self.get_tg_data()[cols]
        df.rename(columns={'提交日期': '日期'}, inplace=True)
        cond = df['小组'].str.endswith('组')
        df = df[cond]
        df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '创量'})
        return df

    def get_tg_teams_groupby(self):
        # 地区 学院  运营部 业绩
        cols = ['地区', '学院', '运营部', '小组']
        df = self.get_tg_data()[cols]
        cols = ['地区', '学院', '运营部']
        cond = df['小组'].str.endswith('组')
        df = df[cond]
        df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '业绩'})
        return df

    def get_tg_groups_groupby_day_evening(self):
        # 地区 学院  运营部 小组 业绩
        hour_line = 18
        df = self.get_tg_data()
        cond1 = ((df['小组'].str.endswith('组')) & (df['提交小时'] < hour_line))
        cond2 = ((df['小组'].str.endswith('组')) & (df['提交小时'] >= hour_line))
        cols = ['地区', '学院', '运营部', '小组', '提交小时']
        df_day = df[cond1][cols]
        df_evening = df[cond2][cols]
        result = [df_day, df_evening]
        result1 = []
        cols = cols[:-1]
        for df in result:
            df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '业绩'})
            result1.append(df)
        return result1

    def get_tg_teams_groupby_day_evening(self):
        # 基于 小组量 -- 地区 学院  运营部 业绩
        result1 = []
        result = self.get_tg_groups_groupby_day_evening()
        for df in result:
            cols_name = list(df.columns)
            df = fun_groupby(df, cols_name[:-2], {cols_name[-1]: 'sum'}, {cols_name[-1]: '业绩'})
            result1.append(df)
        return result1

    def get_tg_all_groupby(self):
        '''
        :return:'地区', '学院', '运营部', '小组', '创量'
        包含运营部、学院、地区的创量
        '''
        cols = ['地区', '学院', '运营部', '小组']
        df = self.get_tg_data()[cols]
        cond = df['小组'].str.endswith('组')
        df = df[cond]
        col_name = '创量'
        df_g = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: col_name})
        df_t = fun_groupby(df_g, cols[:-1], {col_name: 'sum'})
        df_c = fun_groupby(df_t, cols[:-2], {col_name: 'sum'})
        df_r = fun_groupby(df_c, cols[:-3], {col_name: 'sum'})
        df_all = pd.concat([df_g, df_t, df_c, df_r])
        df_all = df_all.reindex(columns=list(df_g.columns))
        return df_all

    def get_tg_peoples_groupby(self):
        # 工号 业绩
        cols = ['工号']
        df = self.get_tg_data()[cols]
        df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '业绩'})
        return df

    def get_tgtj(self):
        '''
        推广统计
        :return:量类型 地区  学院  运营部 小组 业绩 进群量 注册量
        '''
        sql = '''SELECT [量类型],[所属部门] as 地区 ,[所属学院] as 学院 ,[所属战队] as 运营部,[所属分组] as 小组
                ,sum(cast([数据量] as int)) as 业绩 ,sum(cast([进群量] as int)) as 进群量
                ,sum(cast([注册量] as int)) as 注册量
                FROM [QXT].[dbo].[推广统计]
                where [所在岗位] like '推广专员%' and [所属部门] like '{}%'
                group by [量类型],[所属学院],[所属部门],[所属战队],[所属分组]
                having sum(cast([数据量] as int)) <> 0  and sum(cast([进群量] as int)) <> 0
                and sum(cast([注册量] as int)) <> 0'''.format(self.dq)
        df = self.operate_DB(sql)
        df['运营部'] = df['运营部'].map(self.dict_substituet)
        return df

    def get_target(self):
        '''
        原始目标表 运营部
        :return:
        '''
        sql = '''SELECT [地区]
                  ,[学院]
                  ,[战队] [运营部]
                  ,[日目标]
                  ,[月目标]
              FROM [QXT].[dbo].[target]
              WHERE [地区] = '{}' '''.format(self.dq)
        df = self.operate_DB(sql)
        df['运营部'] = df['运营部'].map(self.dict_substituet)
        return df

    def get_holiday(self):
        sql = '''select dt [日期],{} 
                from [dbo].[holidays] 
                where dt between '{}' and '{}'
        '''.format(self.dq, self.start_date, self.final_date)
        return self.operate_DB(sql)

    def get_sum_holidays(self):
        result = self.get_holiday()
        result = result[self.dq].sum()
        return result

    def get_people_num(self):
        '''
        :return: 地区 学院 运营部 小组 日期 人数
        '''
        sql = f'''SELECT [地区]
                  ,[学院]
                  ,[战队] [运营部]
                  ,[小组]
                  ,[日期]
                  ,[人数]
          FROM [QXT].[dbo].[people_num]
          where [地区] = '{self.dq}' and [日期] between '{self.start_date}' and '{self.final_date}' '''
        df = self.operate_DB(sql)
        df['运营部'] = df['运营部'].map(self.dict_substituet)
        return df

    def get_inf_peoplelist(self):
        '''
        所有员工
        :return: '地区', '学院', '运营部', '小组', '员工姓名', '工号',
                '员工岗位', '状态', '入职时间', '离职日期','接量类型', '在职天数'
        '''
        cols = ['地区', '学院', '运营部', '小组', '员工姓名', '工号', '员工岗位', '状态', '入职时间', '离职日期', '接量类型']
        Peoplelist_df = self.get_Peoplelist()[cols]
        Peoplelist_df[cols[-3]] = pd.to_datetime(Peoplelist_df[cols[-3]])
        Peoplelist_df[cols[-2]] = pd.to_datetime(Peoplelist_df[cols[-2]])
        col_name = '在职天数'
        # 在职天数
        final_date_date = MyDate(self.final_date).date_date
        _fun = lambda x: (x[cols[-2]] - x[cols[-3]]) if x[cols[-4]] == '离职' else (
            (final_date_date - x[cols[-3]]) if x[cols[-4]] == '在职' else np.nan)
        Peoplelist_df[col_name] = Peoplelist_df.apply(_fun, axis=1)
        Peoplelist_df[col_name] = np.ceil(Peoplelist_df[col_name] / np.timedelta64(24, 'h')) + 1
        return Peoplelist_df

    def get_inf_group(self):
        '''
        :return: '地区', '学院', '运营部', '小组', '管理员', '工号', '推广人数', '日目标', '月目标'
        '''
        Peoplelist_df = self.get_Peoplelist()
        # 筛选出在职状态 并且 组内人员
        cond = (Peoplelist_df['状态'] == '在职') & (Peoplelist_df['小组'].str.endswith('组'))
        Peoplelist_df = Peoplelist_df[cond]
        # 组内在职人数
        groupby_cols = ['地区', '学院', '运营部', '小组']
        p_num = Peoplelist_df[groupby_cols].groupby(groupby_cols)
        p_num = p_num[groupby_cols[-1]].count()
        p_num = p_num.to_frame()
        p_num.rename({groupby_cols[-1]: '推广人数'}, axis=1, inplace=True)
        # 排序后 去重
        group_df = Peoplelist_df.sort_values(by='员工岗位', ascending=False)
        keep_cols = ['地区', '学院', '运营部', '小组', '员工姓名', '工号', '员工岗位']
        group_df = group_df[keep_cols]
        dup_cils = keep_cols[:4]
        group_df.drop_duplicates(dup_cils, inplace=True)
        # 非组长 组长字段设置为空,
        cond1 = (group_df['员工岗位'] != '推广专员组长')
        group_df['员工姓名'][cond1] = None
        group_df['工号'][cond1] = None
        group_df.drop(keep_cols[-1], axis=1, inplace=True)
        group_df.rename({'员工姓名': '管理员'}, axis=1, inplace=True)
        # 合并 组内人数 与 组信息
        group_df = group_df.merge(p_num, right_index=True, left_on=keep_cols[:4])
        # 得到战队内组个数
        cols2 = keep_cols[:3]
        group_count = group_df[cols2].groupby(cols2)
        group_count = group_count[cols2[2]].count()
        group_count = group_count.to_frame()
        group_count.rename({'运营部': '小组个数'}, axis=1, inplace=True)
        target_df = self.get_target()
        target_df = target_df.merge(group_count, right_index=True, left_on=cols2)
        target_df['小组日目标'] = target_df['日目标'].divide(target_df['小组个数'], fill_value=0)
        target_df['小组月目标'] = target_df['月目标'].divide(target_df['小组个数'], fill_value=0)
        target_df.drop(['日目标', '月目标', '小组个数'], axis=1, inplace=True)
        target_df.rename({'小组日目标': '日目标', '小组月目标': '月目标'}, axis=1, inplace=True)
        # 合并 目标 与 组信息
        result = group_df.merge(target_df, on=cols2)
        return result

    def get_inf_fun(self, df1, s_ind, job_title):
        df = df1.copy()
        e_ind = s_ind + 3
        df_cols = list(df.columns)
        drop_cols = df_cols[s_ind:e_ind]
        df.drop(drop_cols, axis=1, inplace=True)
        groupby_cols = df_cols[:s_ind]
        df = df.groupby(groupby_cols)
        df = df.sum()
        df.reset_index(inplace=True)

        # 得到管理员
        Peoplelist_df = self.get_Peoplelist()
        cond_1 = ((Peoplelist_df['状态'] == '在职') & (Peoplelist_df['员工岗位'] == job_title))
        Peoplelist_df = Peoplelist_df[cond_1]
        cols = groupby_cols + ['员工姓名', '工号']
        leader = Peoplelist_df[cols]
        leader.rename({'员工姓名': '管理员'}, axis=1, inplace=True)

        # 合并 管理人员
        result = pd.merge(df, leader, how='left', on=groupby_cols)
        new_inds = list(range(s_ind)) + [-2, -1] + [-5, -4, -3]
        result = reindex_cols(result, new_inds)
        return result

    def get_inf_team(self):
        '''
        基于 小组信息 得到 运营部信息
        :return:'地区', '学院', '运营部',  '管理员', '工号', '推广人数', '日目标', '月目标'
        '''
        # 其他信息
        df = self.get_inf_group()
        return self.get_inf_fun(df, 3, '推广专员战队长')

    def get_inf_colege(self):
        '''
        基于 运营部信息 得到 学院信息
        :return:'地区', '学院',  '管理员', '工号', '推广人数', '日目标', '月目标'
        '''
        # 其他信息
        df = self.get_inf_team()
        return self.get_inf_fun(df, 2, '学院管理员')

    def get_inf_region(self):
        '''
        基于 运营部信息 得到 学院信息
        :return:'地区', '管理员', '工号', '推广人数', '日目标', '月目标'
        '''
        # 其他信息
        df = self.get_inf_colege()
        return self.get_inf_fun(df, 1, '推广地区管理员')

    def get_inf_all(self):
        g_df = self.get_inf_group()
        t_df = self.get_inf_fun(g_df, 3, '推广专员战队长')
        c_df = self.get_inf_fun(t_df, 2, '学院管理员')
        r_df = self.get_inf_fun(c_df, 1, '推广地区管理员')
        all_df = pd.concat([g_df, t_df, c_df, r_df])
        all_df = all_df.reindex(columns=list(g_df.columns))
        return all_df


class Component():
    def __init__(self, dq, final_date):
        self.dq = dq
        self.final_date = final_date
        # 日期类
        self.cls_final_date = MyDate(final_date)

    # 日报 -- 分量
    def component_df(self):
        # 起始日期 三个月前
        start_date_date = self.cls_final_date.get_date_Nmonthes_firstday(n=-3)
        start_date = get_str_date(start_date_date)
        # 数据类
        cls_sqlserver = MySqlServer(dq=self.dq, start_date=start_date, final_date=self.final_date)
        cols = ['提交日期', '课程顾问-员工号']
        data = cls_sqlserver.get_tg_data_where_guwen()[cols]
        new_cols = ['日期', '工号']
        data.columns = new_cols  # 日期 工号 (课程顾问)

        # 获取每日顾问人数
        unique_consultant = data.drop_duplicates()
        consultant_num = unique_consultant.groupby('日期').count()
        consultant_num.rename(columns={'工号': '顾问人数'}, inplace=True)

        # 获取每日分量
        component_num = data.groupby('日期').count()
        component_num.rename(columns={'工号': '分量'}, inplace=True)

        # 获取日人均分量
        avg_cs_cp = pd.merge(consultant_num, component_num, on='日期')
        avg_cs_cp['日人均分量'] = avg_cs_cp['分量'].divide(avg_cs_cp['顾问人数'], fill_value=0)
        avg_cs_cp['日人均分量'] = avg_cs_cp['日人均分量'].map(lambda x: '%.0f' % x)

        # reindex日期索引列
        new_index = list(pd.date_range(start_date, self.final_date))
        avg_cs_cp = avg_cs_cp.reindex(new_index, fill_value=0)
        avg_cs_cp.reset_index(level=0, inplace=True)
        avg_cs_cp['日期'] = pd.to_datetime(avg_cs_cp.日期)
        avg_cs_cp['月'] = avg_cs_cp.日期.dt.month
        avg_cs_cp['日'] = avg_cs_cp.日期.dt.day

        result = avg_cs_cp[['月', '日', '日人均分量']]
        result.set_index(['月', '日'], inplace=True)
        result = result.unstack()  # 转置
        result.reset_index(inplace=True)
        result['月'] = result['月'].map(lambda x: '%d月' % x)  # 改变index
        result.columns = ['%s日' % col[1] for col in result.columns]  # 改变cols
        result = result.fillna('')  # nan替换为空值

        col_names = list(result.columns)
        col_names[0] = '月'
        result.columns = col_names  # 行名
        return result, start_date, self.final_date


class ReportDataAsDf():
    def __init__(self, dq, final_date):
        self.dq = dq
        self.final_date = final_date

        # 最早产生量的时间
        self.first_date = '2019/6/1'
        # 组员 组长合格线
        self.standard_line = [125, 180]
        self.day_standard = [750, 1000]

        # 日期类
        self.cls_final_date = MyDate(final_date)
        # 终止日期
        self.final_date_date = self.cls_final_date.date_date
        # 本月1号
        self.start_date_date = self.cls_final_date.get_date_Nmonthes_firstday(n=0)
        self.start_date = get_str_date(self.start_date_date)
        # 本月业绩类
        self.cls_sqlserver_0 = MySqlServer(dq=self.dq, start_date=self.start_date, final_date=self.final_date)

        '''
        人员信息:当前数据库中人员表的人员信息
        '地区', '学院', '运营部', '小组', '员工姓名', '工号','员工岗位', '状态', '入职时间', '离职日期','接量类型', '在职天数'
        '''
        self.inf_people = self.cls_sqlserver_0.get_inf_peoplelist()  # 所有人员
        self.inf_people_columns = list(self.inf_people.columns)
        self.inf_groupleader = self.groupleader_inf(self.inf_people)  # 组长信息 与所有人员字段一致

        self.tg_peoples_groupby_0 = self.cls_sqlserver_0.get_tg_peoples_groupby()  # 本月员工业绩  工号 业绩

        '''
        组信息
        '地区', '学院', '运营部', '小组', '管理员', '工号', '推广人数', '日目标', '月目标'
        '''
        self.inf_group = self.cls_sqlserver_0.get_inf_group()
        self.inf_group_columns = list(self.inf_group.columns)
        self.group_day_df, self.group_evening_df = self.cls_sqlserver_0.get_tg_groups_groupby_day_evening()  # 白天 和 夜间业绩  '地区', '学院', '运营部' 小组 业绩
        '''
        运营部信息
        '地区', '学院', '运营部', '管理员', '工号', '推广人数', '日目标', '月目标'
        '''
        self.inf_team = self.cls_sqlserver_0.get_inf_team()
        self.inf_team_columns = list(self.inf_team.columns)
        self.team_day_df, self.team_evening_df = self.cls_sqlserver_0.get_tg_teams_groupby_day_evening()  # 白天 和 夜间业绩  '地区', '学院', '运营部' 业绩
        # 本月工作天数
        self.sum_workday = self.cls_sqlserver_0.get_sum_holidays()

    # 组长信息
    def groupleader_inf(self, people_df):
        cond0 = (people_df['员工岗位'] == '推广专员组长')
        cond1 = (people_df['状态'] == '在职')
        return people_df[cond0 & cond1]

    def erollment_and_group_entry_rate_1(self, data, cols):
        # erollment_and_group_entry_rate_df的辅助函数
        other_cols = ['业绩', '注册量', '注册率', '进群量', '进群率']
        final_data = data.groupby(cols)
        final_data = final_data.sum()
        final_data.reset_index(inplace=True)
        final_data[other_cols[-1]] = final_data[other_cols[-2]].divide(final_data[other_cols[0]], fill_value=0).replace(
            [np.nan, np.inf], 0)  # 进群率
        final_data[other_cols[2]] = final_data[other_cols[1]].divide(final_data[other_cols[0]], fill_value=0).replace(
            [np.nan, np.inf], 0)  # 注册率

        cls_format = FormatSeriesOfDf(final_data)
        final_data = cls_format.format_dot(cols_name=(other_cols[0], other_cols[1], other_cols[-2]))
        final_data = cls_format.format_percentage(cols_name=(other_cols[2], other_cols[-1]))

        # final_data[other_cols[0]] = final_data[other_cols[0]].map(lambda x: '%.0f' % x)  # 业绩
        # final_data[other_cols[1]] = final_data[other_cols[1]].map(lambda x: '%.0f' % x)
        # final_data[other_cols[-2]] = final_data[other_cols[-2]].map(lambda x: '%.0f' % x)
        # final_data[other_cols[2]] = final_data[other_cols[2]].map(lambda x: '%.0f%s' % (x * 100, '%'))
        # final_data[other_cols[-1]] = final_data[other_cols[-1]].map(lambda x: '%.0f%s' % (x * 100, '%'))
        new_cols = cols.extend(other_cols)
        final_data = final_data.reindex(new_cols, axis=1)
        return final_data

    # 日报 -- 注册率与进群率
    def erollment_and_group_entry_rate_df(self):
        # 获取推广统计数据
        data = self.cls_sqlserver_0.get_tgtj()

        # 替换地区 部门 推广一部-'' 1战队-推广一部
        data['地区'].replace(r'推广\w部', '', regex=True, inplace=True)
        cols = list(data.columns)  # 量类型 地区  学院  运营部 小组 业绩 进群量 注册量
        deparment_list = cols[:5]
        group_data, team_data, colege_data, region_data = [
            self.erollment_and_group_entry_rate_1(data, deparment_list[:_]) for _ in range(5, 1, -1)]
        return group_data, team_data, colege_data, region_data, self.start_date

    # 日报 -- 近4月每日业绩趋势
    def trend_4monthes_df(self):
        # 获取小组业绩
        start_date = get_str_date(self.cls_final_date.get_date_Nmonthes_firstday(n=-3))
        data = MySqlServer(dq=self.dq, start_date=start_date, final_date=self.final_date).get_tg_groups_groupby_date()
        cols = list(data.columns)  # '地区', '学院', '运营部', '小组', '日期', '创量'
        # 地区
        dep_colege = cols[0:1]
        data_region_list = self.trend_4monthes_1(data, dep_colege, start_date)

        # 学院
        dep_region = cols[0:2]
        data_colege_list = self.trend_4monthes_1(data, dep_region, start_date)

        # 运营部
        dep_team = cols[0:3]
        data_team_list = self.trend_4monthes_1(data, dep_team, start_date)

        return data_region_list, data_colege_list, data_team_list, start_date

    def trend_4monthes_1(self, df, cols, start_date):
        # trend_4monthes_df的辅助函数
        # 获取列表集
        cols_index = df[cols].drop_duplicates()
        col_name1 = '日期'
        date_index = list(pd.date_range(start_date, self.final_date))
        date_index = pd.DataFrame(date_index, columns=[col_name1])
        # 获取笛卡尔索引
        new_index = cartesian_product_df(cols_index, date_index)
        # 聚合
        col_names = list(df.columns)
        cols.append(col_name1)  # 组织架构  日期
        df_groupby = df.groupby(cols)
        df_sum = df_groupby[col_names[-1]].sum()
        # 重新设置索引
        df_sum = df_sum.reindex(new_index, fill_value=0)
        df_sum = df_sum.to_frame()
        df_sum.reset_index(inplace=True)
        df_sum[col_name1] = pd.to_datetime(df_sum[col_name1])

        # 获取本月工作日平均量
        df_holiday = self.cls_sqlserver_0.get_holiday()
        df_holiday[col_name1] = pd.to_datetime(df_holiday[col_name1])
        df_holiday = df_holiday[[col_name1, self.dq]]
        df_avg = pd.merge(df_sum, df_holiday, how='left', on=col_name1)
        df_avg = df_avg[(df_avg[self.dq] == 1) & (df_avg[col_names[-1]] != 0)]
        df_avg = df_avg.groupby(cols[:-1])
        df_avg = df_avg[col_names[-1]].mean()

        col_names1 = ['月', '日'] + col_names[-1:]
        df_sum[col_names1[0]] = df_sum.日期.dt.month
        df_sum[col_names1[1]] = df_sum.日期.dt.day

        cols2 = cols[:-1] + col_names1
        result = df_sum[cols2]
        result.set_index(cols2[:-1], inplace=True)

        result = result.unstack()  # 转置
        result.reset_index(inplace=True)
        result['月'] = result['月'].map(lambda x: '%d月' % x)  # 改变index
        result = result.fillna('')  # nan替换为空值
        col_names = list(result.columns)
        for ind, col in enumerate(col_names):
            if not col[1]:
                col_names[ind] = col[0]
            else:
                col_names[ind] = '{}日'.format(col[1])
        result.columns = col_names  # 行名
        result_list = self.trend_4monthes_2(result, cols[:-1], df_avg)  # 唯一化输出到列表
        return result_list

    def trend_4monthes_2(self, df, cols, df_avg):
        # trend_4monthes_df的辅助函数
        # 获取每种架构一个列表集
        result = []
        len_cols = len(cols)
        indexs = df[cols].drop_duplicates()
        for index in range(indexs.shape[0]):
            new_df = indexs.iloc[index]
            new_df = new_df.to_frame().T
            new_df = pd.merge(new_df, df, on=cols)

            avg_row = new_df.iloc[-1].copy()
            avg_row['月'] = avg_row['月'] + '工作日日均量'

            # 获取工作日平均量
            avg_row = avg_row.to_frame().T
            if df_avg.empty:
                avg = 0
            else:
                avg = pd.merge(df_avg, avg_row, left_index=True, right_on=cols)['创量']
            avg_row.iloc[:, (len_cols + 1):] = avg
            new_df = pd.concat([new_df, avg_row])
            for colname in list(new_df.columns)[(len_cols + 1):]:
                new_df[colname] = new_df[colname].map(lambda x: '%.0f' % x if not isinstance(x, str) else '')
            result.append(new_df)
        return result

    # 日报-- 与上月最高业绩对比
    def vs_last_month_df(self):
        cols = self.inf_team_columns[:3]  # '学院', '地区', '运营部'
        inf_team = self.inf_team[cols]

        # 获取运营部本月业绩 和 均量
        data = self.vs_last_month_1(self.dq, self.start_date, self.final_date, inf_df=inf_team, cols=cols)

        # 获取上月业绩
        start_date_1 = get_str_date(self.cls_final_date.get_date_Nmonthes_firstday(-1))
        final_date_1 = get_str_date(self.cls_final_date.get_date_Nmonth_sameday(-1))
        data_1 = self.vs_last_month_1(self.dq, start_date_1, final_date_1, inf_df=inf_team, cols=cols)

        result = pd.merge(data_1, data, how='left', on=cols)
        new_list_index = [0, 1, 2, 5, 9, 6, 10, 3, 4, 7, 8]
        result = reindex_cols(result, new_list_index)
        return result, self.start_date

    def vs_last_month_1(self, dq, start_date, final_date, inf_df, cols):
        cls_sqlserver = MySqlServer(dq=dq, start_date=start_date, final_date=final_date)
        data = cls_sqlserver.get_tg_teams_groupby()
        col = get_str_date(final_date, '%m') + '月'

        col_amount = col + '业绩'
        data = data.rename({'业绩': col_amount}, axis=1)

        # 获取工作天数
        sum_workday = cls_sqlserver.get_sum_holidays()
        col_workday = col + '工作天数'
        data[col_workday] = sum_workday

        # 获取日均业绩
        col_day_avg = col + '日均业绩'
        data[col_day_avg] = data[col_amount].div(sum_workday, fill_value=0).replace(np.inf, 0).replace(np.nan, 0)

        # 获取全部运营部
        data = pd.merge(inf_df, data, on=cols, how='left')

        # 获取日均业绩均线
        col_day_avg_line = col + '日均业绩均线'
        data[col_day_avg_line] = data[col_day_avg].mean()

        data = data.fillna(0)

        change_cols = [col_amount, col_workday, col_day_avg, col_day_avg_line]
        data = FormatSeriesOfDf(df=data).format_dot(n=0, cols_name=change_cols)
        return data

    # 日报 -- 完成率与时间消耗率
    def complete_rate_time_rate_df(self):
        cols = self.inf_team_columns[:3] + self.inf_team_columns[-1:]  # '地区', '学院', '运营部', '月目标'
        # 获取运营部及月目标
        target_team = self.inf_team[cols]

        # 获取inf_cols
        inf_cols = cols[:-1]

        # 获取本月业绩
        data = self.cls_sqlserver_0.get_tg_teams_groupby()
        col = get_str_date(self.final_date, '%m') + '月'
        col_amount = col + '业绩'
        data = data.rename({'业绩': col_amount}, axis=1)

        # 获取完成率
        col_name = '完成率'
        result = pd.merge(target_team, data, on=inf_cols, how='left')
        result[col_name] = result[col_amount].divide(result[cols[-1]], fill_value=0).replace(np.inf, 0).replace(np.nan,
                                                                                                                0)

        # 工作天数
        col_workday = col + '工作天数'
        result[col_workday] = self.sum_workday

        # 获取本月总的工作天数
        final_date_0 = get_str_date(self.cls_final_date.get_date_Nmonthes_endday())
        sum_workday_1 = MySqlServer(dq=self.dq, start_date=self.start_date, final_date=final_date_0).get_sum_holidays()
        col_workday_1 = col + '总工作天数'
        result[col_workday_1] = sum_workday_1

        # 获取时间消耗率
        col_complete_rate = '时间消耗率'
        result[col_complete_rate] = result[col_workday].divide(result[col_workday_1], fill_value=0).replace(np.inf,
                                                                                                            0).replace(
            np.nan, 0)

        result = FormatSeriesOfDf(df=result).format_percentage(n=0, cols_name=[col_complete_rate, col_name])

        # 改变列顺序
        new_list_index = [0, 1, 2, 5, 8, 3, 4, 6, 7]
        result = reindex_cols(result, new_list_index)
        return result, self.start_date

    # 日报 -- 组长业绩贡献率
    def group_leader_rate(self):
        cols = self.inf_group_columns[:6]  # '地区', '学院', '运营部', '小组', '管理员', '工号'
        inf = self.inf_group[cols]  # 获取组长信息

        # 获取本月组内人员业绩
        groupby_people = self.tg_peoples_groupby_0  # 工号 业绩
        # 获取本月组长业绩
        col_names = ['组长业绩', '组长']
        result = pd.merge(inf, groupby_people, on=cols[-1], how='left')
        result = result.rename({'业绩': col_names[0], '管理员': col_names[1]},
                               axis=1)  # '地区', '学院', '运营部', '小组', '组长', '工号' 组长业绩

        # 获取本月小组业绩
        groups_cols = cols[:4]
        groups_df = self.cls_sqlserver_0.get_tg_groups_groupby()
        result = pd.merge(result, groups_df, on=groups_cols, how='left')
        col_names1 = ['小组业绩', '组长贡献率']
        result = result.rename({'创量': col_names1[0]}, axis=1)
        result[[col_names[0], col_names1[0]]] = result[[col_names[0], col_names1[0]]].fillna(0)
        result[col_names1[1]] = result[col_names[0]].divide(result[col_names1[0]], fill_value=0).replace(np.inf,
                                                                                                         0).replace(
            np.nan, 0)

        result = FormatSeriesOfDf(df=result).format_percentage(n=0, cols_name=col_names1[1])

        # 改变列顺序
        new_list_index = [0, 1, 2, 3, 6, 7, 8, 4, 5]
        result = reindex_cols(result, new_list_index)
        result.sort_values(by='小组业绩', ascending=False, inplace=True)
        return result, self.start_date

    # 日报-- 运营部白天夜间业绩对比
    def team_day_evening(self):
        cols = self.inf_team_columns[:3]
        inf_team = self.inf_team[cols]  # '地区', '学院', '运营部'
        col_names0 = ['白天业绩', '夜间业绩', '总业绩', '工作天数', '日均业绩', '白日占比', '业绩均线']
        self.team_day_df.rename({'业绩': col_names0[0]}, axis=1, inplace=True)
        self.team_evening_df.rename({'业绩': col_names0[1]}, axis=1, inplace=True)
        data = pd.merge(self.team_day_df, self.team_evening_df, how='outer', on=cols)
        data[col_names0[2]] = data[col_names0[0]] + data[col_names0[1]]
        avg_teams = data[col_names0[2]].mean()

        # 工作天数
        data[col_names0[3]] = self.sum_workday
        # 获取日均业绩
        data[col_names0[4]] = data[col_names0[2]].div(self.sum_workday).replace(np.inf, 0).replace(np.nan, 0)
        # 获取白日占比
        data[col_names0[5]] = data[col_names0[0]].divide(data[col_names0[2]]).replace(np.inf, 0).replace(np.nan, 0)
        # 获取全部运营部
        data = pd.merge(inf_team, data, on=cols, how='left')
        # 获取均线
        data[col_names0[6]] = avg_teams
        data = data.fillna(0)
        # 格式
        cls_format = FormatSeriesOfDf(df=data)
        data = cls_format.format_dot(cols_name=[col_names0[4], col_names0[6]])
        data = cls_format.format_percentage(cols_name=col_names0[5])
        new_list_index = [0, 1, 2, 4, 3, 5, 7, 9, 8, 6]
        result = reindex_cols(data, new_list_index)
        return result, self.start_date

    # 日报-- 运营部 - 推广小组白天夜间业绩对比
    def group_team_day_evening(self):
        cols_g = self.inf_group_columns[:4]
        inf_g = self.inf_group[cols_g]  # '地区', '学院', '运营部','小组'
        dfs_g = [self.group_day_df, self.group_evening_df]

        cols_t = self.inf_team_columns[:3]
        inf_t = self.inf_team[cols_t]  # '地区', '学院', '运营部'
        dfs_t = [self.team_day_df, self.team_evening_df]

        cols_all = [cols_g, cols_t]
        inf_all = [inf_g, inf_t]
        dfs_all = [dfs_g, dfs_t]
        index_all = [[0, 1, 2, 3, 5, 4, 6, 8, 10, 9, 7], [0, 1, 2, 4, 3, 5, 7, 9, 8, 6]]
        result_all = []
        for i in range(2):
            col_names0 = ['白天业绩', '夜间业绩', '总业绩', '工作天数', '日均业绩', '白日占比', '业绩均线']
            dfs_all[i][0].rename({'业绩': col_names0[0]}, axis=1, inplace=True)
            dfs_all[i][1].rename({'业绩': col_names0[1]}, axis=1, inplace=True)
            data = pd.merge(dfs_all[i][0], dfs_all[i][1], how='outer', on=cols_all[i])
            data[col_names0[2]] = data[col_names0[0]] + data[col_names0[1]]
            avg_line = data[col_names0[2]].mean()

            # 工作天数
            data[col_names0[3]] = self.sum_workday
            # 获取日均业绩
            data[col_names0[4]] = data[col_names0[2]].div(self.sum_workday).replace([np.inf, np.nan], 0)
            # 获取白日占比
            data[col_names0[5]] = data[col_names0[0]].divide(data[col_names0[2]]).replace([np.inf, np.nan], 0)
            # 获取全部组织架构
            data = pd.merge(inf_all[i], data, on=cols_all[i], how='left')
            # 获取均线
            data[col_names0[6]] = avg_line
            data = data.fillna(0)
            # 格式
            cls_format = FormatSeriesOfDf(df=data)
            data = cls_format.format_dot(cols_name=[col_names0[4], col_names0[6]])
            data = cls_format.format_percentage(cols_name=col_names0[5])
            data.sort_values(by=col_names0[2], ascending=False, inplace=True)

            new_list_index = index_all[i]
            result = reindex_cols(data, new_list_index)
            result_all.append(result)
        result_all.append(self.start_date)
        return result_all

    # 日报-- 近6个月业绩对比
    def team_six_month(self):
        # 获取小组
        cols = self.inf_group_columns[:4]  # '地区', '学院', '运营部', '小组'
        inf = self.inf_group[cols]
        # 获取全部索引
        col_name0 = '日期'
        start_date_5 = self.cls_final_date.get_date_Nmonthes_firstday(n=-5)
        date_index = list(pd.date_range(start_date_5, self.final_date))
        date_index = pd.DataFrame(date_index, columns=[col_name0])
        cartesian_df = cartesian_product_df(inf, date_index)
        # 筛选符合要求的索引
        cols.append(col_name0)  # '地区', '学院', '运营部', '小组', '日期'
        cartesian_df['月'] = cartesian_df[col_name0].dt.month
        cartesian_df['日'] = cartesian_df[col_name0].dt.day
        current_day = int(get_str_date(self.final_date, '%d'))
        col_name = '列名'
        month_list = [self.cls_final_date.get_date_Nmonth_sameday(i) for i in range(-5, 1)]
        col_dict = {dt.month: f'{dt.month}月1日-{dt.month}月{dt.day}日' for dt in month_list}
        cartesian_df[col_name] = cartesian_df['月'].map(col_dict)
        cond1 = cartesian_df['日'] <= current_day
        cartesian_df = cartesian_df[cond1][cols + [col_name]]
        # 获取小组近6月同期业绩
        cls_mysqlserver = MySqlServer(dq=self.dq, start_date=start_date_5, final_date=self.final_date)
        group_df = cls_mysqlserver.get_tg_groups_groupby_date()
        group_df[col_name0] = pd.to_datetime(group_df[col_name0])
        result_g = pd.merge(cartesian_df, group_df, on=cols, how='left')
        col_name2 = result_g.columns[-1]  # 创量
        result_g[col_name2] = result_g[col_name2].fillna(0)

        # 获取小组近6月同期除休假外的工作总人数
        people_num_df = cls_mysqlserver.get_people_num()
        holiday_df = cls_mysqlserver.get_holiday()
        people_num_df = pd.merge(people_num_df, holiday_df, on=[col_name0])
        cond2 = (people_num_df[self.dq] == 1)
        people_num_df = people_num_df[cond2][people_num_df.columns[:-1]]
        people_num_df[col_name0] = pd.to_datetime(people_num_df[col_name0])

        # 合并 业绩 和 日期段内在职总人数
        result_group = pd.merge(result_g, people_num_df, how='left', on=cols)
        col_name3 = list(result_group.columns)[-1]
        result_group[col_name3] = result_group[col_name3].fillna(0)
        del result_group[col_name0]
        cols.pop(-1)
        cols.append(col_name)  # '地区', '学院', '运营部', '小组', '列名'
        result_group = fun_groupby(result_group, cols, {col_name2: 'sum', col_name3: 'sum'})

        # 运营部
        result_team = result_group.copy()
        del result_team[cols[3]]
        team_cols = list(result_team.columns)
        groupby_cols = team_cols[:4]
        cols_fun = {k: 'sum' for k in team_cols[-2:]}
        result_team = fun_groupby(result_team, groupby_cols, cols_fun)

        # 学院
        result_colege = result_team.copy()
        del result_colege[cols[2]]
        colege_cols = list(result_colege.columns)
        groupby_cols = colege_cols[:3]
        result_colege = fun_groupby(result_colege, groupby_cols, cols_fun)

        # 地区
        result_region = result_colege.copy()
        del result_region[cols[1]]
        region_cols = list(result_region.columns)
        groupby_cols = region_cols[:2]
        result_region = fun_groupby(result_region, groupby_cols, cols_fun)

        # 返回列表
        all_result = [self.team_six_month_fun(_) for _ in [result_group, result_team, result_colege, result_region]]
        all_result.append(self.start_date)
        return all_result

    def team_six_month_fun(self, result):
        result_columns0 = list(result.columns)
        col_name = '日人均创量'
        result[col_name] = result[result_columns0[-2]].divide(result[result_columns0[-1]], fill_value=0).replace(np.inf,
                                                                                                                 0).replace(
            np.nan, 0)
        result = FormatSeriesOfDf(result).format_dot(cols_name=col_name)
        result = result.set_index(list(result.columns)[:-3])
        result = result.stack()
        result = result.unstack(-2)
        result = result.reset_index()
        result_columns1 = list(result.columns)
        cols_len = len(result_columns1)
        new_indexs = list(range(cols_len - 6, cols_len))
        new_indexs = [cols_len - 7] + list(range((cols_len - 7))) + new_indexs
        result = reindex_cols(result, new_indexs)
        result.sort_values(by=list(result.columns)[:-6], inplace=True)
        result_columns2 = list(result.columns)
        result_columns2[0] = '类型'
        result.columns = result_columns2
        return result

    # 日报-- 运营部夜间业绩对比 包含小组 和 运营部
    def evening(self):
        cols = [self.inf_group_columns[:4], self.inf_team_columns[:3]]
        infs = [self.inf_group, self.inf_team]
        dfs = [self.group_evening_df, self.team_evening_df]
        result = []
        for ind, df in enumerate(dfs):
            col_name = '夜间业绩'
            new_cols = cols[ind]
            inf = infs[ind][new_cols]
            # 获取本月业绩
            evening_df = deepcopy(df)
            evening_df.rename({'业绩': col_name}, axis=1, inplace=True)
            avg_teams = evening_df[col_name].mean()
            # 获取全部运营部
            data = pd.merge(inf, evening_df, on=new_cols, how='left')
            # 获取均线
            data['业绩均线'] = round(avg_teams)
            data = data.fillna(0)
            data.sort_values(by=col_name, ascending=False, inplace=True)
            result.append(data)
        result.append(self.start_date)
        return result

    # 日报 -- 运营部人员信息
    def team_peoplelist(self):
        start_date_date = self.start_date_date
        final_date_date = self.final_date_date

        l1_start_date = self.cls_final_date.get_date_Nmonthes_firstday(-1)
        l1_final_date = self.cls_final_date.get_date_Nmonth_sameday(-1)

        df = self.inf_people
        cond1 = df['小组'].notnull()
        df = df[cond1]
        cond = df['小组'].str.endswith('组')
        cols = ['地区', '学院', '运营部', '小组', '状态', '入职时间', '离职日期', '在职天数']
        df = df[cond][cols]
        columns = cols[:4] + ['在职人数', '在职人数(在职天数>180)', '在职人数(在职天数 91-180)', '在职人数(在职天数 61-90)',
                              '在职人数(在职天数 31-60)', '在职人数(在职天数 7-30)', '在职人数(在职天数 1-6)',
                              '离职人数(在职天数>180)', '离职人数(在职天数 91-180)', '离职人数(在职天数 61-90)',
                              '离职人数(在职天数 31-60)', '离职人数(在职天数 7-30)', '离职人数(在职天数 1-6)',
                              f'入职人数（{self.final_date}）', '本月入职人数', '本月离职人数', '本月新进员工离职人数',
                              '本月入职-本月离职', '上月同期离职人数', '上月同期在职人数',
                              '本月新人离职率', '本月离职率', '上月同期离职率']
        result = df.copy()
        result = result.reindex(columns=columns, fill_value=0)
        for i in range(len(df)):
            if df.iloc[i, 5] == final_date_date:  # 当天入职
                result.iloc[i, 17] = 1

            if df.iloc[i, 4] == '在职':
                result.iloc[i, 4] = 1
                if df.iloc[i, 5] <= l1_final_date:
                    result.iloc[i, 23] = 1
                elif start_date_date <= df.iloc[i, 5] <= final_date_date:
                    result.iloc[i, 18] = 1

                if df.iloc[i, 7] > 180:
                    result.iloc[i, 5] = 1
                elif 91 <= df.iloc[i, 7] <= 180:
                    result.iloc[i, 6] = 1
                elif 61 <= df.iloc[i, 7] <= 90:
                    result.iloc[i, 7] = 1
                elif 31 <= df.iloc[i, 7] <= 60:
                    result.iloc[i, 8] = 1
                elif 7 <= df.iloc[i, 7] <= 30:
                    result.iloc[i, 9] = 1
                elif 1 <= df.iloc[i, 7] <= 6:
                    result.iloc[i, 10] = 1

            elif df.iloc[i, 4] == '离职':
                if df.iloc[i, 5] <= l1_final_date and df.iloc[i, 6] > l1_final_date:  # 上月同期末前入职 后离职
                    result.iloc[i, 23] = 1

                if l1_start_date <= df.iloc[i, 6] <= l1_final_date:  # 上月同期离职
                    result.iloc[i, 22] = 1
                elif start_date_date <= df.iloc[i, 6] <= final_date_date:  # 本月离职
                    result.iloc[i, 19] = 1
                    if start_date_date <= df.iloc[i, 5] <= final_date_date:  # 本月入职
                        result.iloc[i, 20] = 1
                        result.iloc[i, 18] = 1

                    if df.iloc[i, 7] > 180:
                        result.iloc[i, 11] = 1
                    elif 91 <= df.iloc[i, 7] <= 180:
                        result.iloc[i, 12] = 1
                    elif 61 <= df.iloc[i, 7] <= 90:
                        result.iloc[i, 13] = 1
                    elif 31 <= df.iloc[i, 7] <= 60:
                        result.iloc[i, 14] = 1
                    elif 7 <= df.iloc[i, 7] <= 30:
                        result.iloc[i, 15] = 1
                    elif 1 <= df.iloc[i, 7] <= 6:
                        result.iloc[i, 16] = 1

        result_group = fun_groupby(result, columns[:4], {columns[_]: 'sum' for _ in range(4, 27)})
        cond2 = (result_group[columns[4]] != 0)
        result_group = result_group[cond2]

        result_team = result_group.copy()
        del result_team[columns[3]]
        result_team = fun_groupby(result_team, columns[:3], {columns[_]: 'sum' for _ in range(4, 27)})

        result_colege = result_team.copy()
        del result_colege[columns[2]]
        result_colege = fun_groupby(result_colege, columns[:2], {columns[_]: 'sum' for _ in range(4, 27)})

        result_region = result_colege.copy()
        del result_region[columns[1]]
        result_region = fun_groupby(result_region, columns[:1], {columns[_]: 'sum' for _ in range(4, 27)})

        r_ls = [result_group, result_team, result_colege, result_region]

        for r in r_ls:
            r[columns[-6]] = r[columns[-9]] - r[columns[-8]]  # 本月入职-本月离职
            r[columns[-3]] = r[columns[-7]].divide(r[columns[-9]], fill_value=0).replace([np.inf, np.nan], 0)  # 当月新人离职率
            r[columns[-2]] = r[columns[-8]].divide(r[columns[-8]] + r[columns[4]], fill_value=0).replace(
                [np.inf, np.nan], 0)  # '本月离职率'
            r[columns[-1]] = r[columns[-5]].divide(r[columns[-5]] + r[columns[-4]], fill_value=0).replace(
                [np.inf, np.nan], 0)  # '上月同期离职率'
            r = FormatSeriesOfDf(r).format_percentage(cols_name=columns[-3:])
            del r[columns[-5]]
            del r[columns[-4]]

        result_group, result_team, result_colege, result_region = r_ls
        # 合并
        result_all = pd.concat(r_ls, ignore_index=True)
        result_all = result_all.reindex(columns=list(result_group.columns), fill_value='')
        result_all.sort_values(by=columns[:4], inplace=True)
        # 填充总计
        result_all = fill_na(result_all, 3, [2, 1, 0])
        # 填充不必要的单元格
        result_all[columns[:3]] = result_all[columns[:3]].fillna('-')
        return result_all, self.start_date

    # 日报 -- 目标完成率
    def complete_rate_df(self):
        cols = ['地区', '学院', '运营部', '小组', '月目标']
        # 获取小组及以上架构和目标
        inf_all = self.cls_sqlserver_0.get_inf_all()[cols]
        # 获取本月所有架构业绩
        data = self.cls_sqlserver_0.get_tg_all_groupby()
        col_name3 = list(data.columns)[-1]
        col_name = '本月业绩'
        data.rename(columns={col_name3: col_name}, inplace=True)
        # 合并
        result = pd.merge(left=inf_all, right=data, how='left', on=cols[:4])
        # 填充总计-
        result = fill_na(result, 3, [2, 1, 0])
        # 添加辅助列 排序
        col_name2 = 'rank'
        result[col_name2] = result[cols[3]].map(lambda x: 1 if x.endswith('组') else 2)
        # 排序
        result.sort_values(by=[col_name], ascending=False, inplace=True)
        result.sort_values(by=cols[:3] + [col_name2], inplace=True)
        result[cols[:3]] = result[cols[:3]].fillna('-')
        result.drop(col_name2, axis=1, inplace=True)
        # 距目标差值
        col_name1 = '距目标差值'
        result[col_name1] = result[col_name] - result[cols[-1]]
        col_name2 = '完成率'
        result[col_name2] = format_series_percentage(
            result[col_name].divide(result[cols[-1]]).replace([np.nan, np.inf], 0))
        # 小数点后0位
        columns = list(result.columns)
        for _ in [4, 6]:
            result[columns[_]] = result[columns[_]].map(lambda x: format(x, '.0f'))
        return result, self.start_date

    # 日报 -- 组长组员日人均业绩
    def group_leader_member(self):
        # '地区', '学院', '运营部', '小组', '员工姓名', '工号', '员工岗位', '状态', '入职时间', '离职日期', '接量类型', '在职天数'
        inds = [0, 1, 2, 3, 6, 5, 7, 11]
        cols = [self.inf_people_columns[_] for _ in inds]  # '地区', '学院', '运营部', '小组', '员工岗位', '工号', '状态', '在职天数'
        # cols = ['地区', '学院', '运营部', '小组', '员工岗位', '工号', '状态', '在职天数']
        # 获取满1月在职员工信息

        inf_p = self.inf_people[cols]
        cond0 = inf_p[cols[3]].notnull()
        cond1 = inf_p[cols[-1]] > 29
        cond2 = inf_p[cols[3]].str.endswith('组')
        cond3 = inf_p[cols[-2]] == '在职'
        inf_p = inf_p[cond0 & cond1 & cond2 & cond3]

        # 获取本月组内人员业绩
        groupby_p = self.tg_peoples_groupby_0

        # 获取本月工作天数
        onwork_days = self.sum_workday

        # 合并
        result = pd.merge(inf_p, groupby_p, on=cols[-3], how='left')
        col_name = '业绩'
        result[col_name] = result[col_name].fillna(0)
        result[col_name] = result[col_name].divide(onwork_days, fill_value=0).replace([np.inf, np.nan], 0)
        result.drop(cols[-3:], axis=1, inplace=True)
        result = fun_groupby(result, groupby_cols=cols[:5], cols_fun_dict={col_name: 'mean'})

        # 行转列
        result.set_index(cols[:5], inplace=True)
        # 转置
        result = result.unstack()
        result = result.fillna(0)
        result.reset_index(inplace=True)
        result.columns = [col[1] if col[1] else col[0] for col in result.columns]
        col_names0 = ['在职满1月组员日均业绩', '组长日均业绩']
        result.rename(columns={'推广专员': col_names0[0], '推广专员组长': col_names0[1]}, inplace=True)
        for _ in col_names0:
            result[_] = result[_].round()
        # 合格线
        col_names1 = ['组员日均量合格线', '组长日均量合格线']
        for ind, colname in enumerate(col_names1):
            result[colname] = self.standard_line[ind]
        result.sort_values(cols[:4], inplace=True)
        return result, self.start_date

    # 日报 -- 单日业绩统计
    def people_performance(self):
        # 获取信息
        col_names = ['白日业绩', '晚间业绩', '合计']
        cols = ['地区', '学院', '小组', '账号类型', '工号', '推广专员', '提交小时']
        cls_sqlserver_1 = MySqlServer(dq=self.dq, start_date=self.final_date, final_date=self.final_date)
        df = cls_sqlserver_1.get_tg_data()[cols]
        cond = df[cols[2]].str.endswith('组')
        df = df[cond]

        df[cols[-1]] = df[cols[-1]].map(lambda x: col_names[0] if x < 18 else col_names[1])
        df.drop(cols[2], axis=1, inplace=True)
        cols.pop(2)

        df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '业绩'})
        # 转置
        df.set_index(cols[:], inplace=True)
        result = df.unstack()
        result = result.fillna(0)
        result.reset_index(inplace=True)
        result.columns = [col[1] if col[1] else col[0] for col in result.columns]
        # 合计
        for _ in col_names[:2]:
            if _ not in result.columns:
                result[_] = 0
        result[col_names[2]] = result[col_names[0]] + result[col_names[1]]
        # 排序
        result.sort_values(col_names[2], ascending=False, inplace=True)
        result = result.reindex(columns=list(result.columns)[:-3] + col_names)
        # result.columns = list(result.columns)[:-3] + col_names
        # 排名
        result.insert(0, '排名', list(range(1, result.shape[0] + 1)))
        return result, self.start_date

    # 日报 -- 推广专员月内日均创量排名
    def people_month_avg_rank(self):
        cols = ['工号', '员工姓名', '地区', '学院', '运营部', '小组', '员工岗位', '入职时间', '状态', '在职天数']
        # 在职员工信息
        inf_p = self.inf_people[cols]
        # 获取组内员工
        cond0 = inf_p[cols[5]].notnull()
        cond2 = inf_p[cols[5]].str.endswith('组')
        cond3 = inf_p[cols[-2]] == '在职'
        inf_p = inf_p[cond0 & cond2 & cond3]
        # 删除状态
        inf_p.drop(cols[-2], axis=1, inplace=True)
        cols.pop(-2)  # '工号', '员工姓名', '地区', '学院', '运营部', '小组', '员工岗位', '入职时间', '在职天数'

        # 月内工作天数
        col_name0 = '月内工作天数'
        inf_p[col_name0] = inf_p[cols[-2]].map(
            lambda x: self.sum_workday if x < self.start_date_date else MySqlServer(dq=self.dq,
                                                                                    start_date=x,
                                                                                    final_date=self.final_date).get_sum_holidays())
        cols.append(col_name0)  # '工号', '员工姓名', '地区', '学院', '运营部', '小组', '员工岗位', '入职时间', '在职天数','月内工作天数'

        # 得到本月业绩
        col_name3 = '本月业绩'
        tg_this_month = self.tg_peoples_groupby_0
        tg_this_month.rename(columns={'业绩': col_name3}, inplace=True)
        result = pd.merge(inf_p, tg_this_month, on=cols[0], how='left')
        result[col_name3].replace(np.nan, 0, inplace=True)
        cols.append(col_name3)  # '工号', '员工姓名', '地区', '学院', '运营部', '小组', '员工岗位', '入职时间', '在职天数','月内工作天数','本月业绩'

        # 得到之前月份业绩
        col_name2 = '之前月份业绩'
        last_1m_endd_date = self.cls_final_date.get_date_Nmonthes_endday(n=-1)
        tg_other_month = MySqlServer(dq=self.dq, start_date=self.first_date,
                                     final_date=last_1m_endd_date).get_tg_peoples_groupby()
        tg_other_month.rename(columns={'业绩': col_name2}, inplace=True)
        result = pd.merge(result, tg_other_month, on=cols[0], how='left')
        result[col_name2].replace(np.nan, 0, inplace=True)
        cols.append(col_name2)  # '工号', '员工姓名', '地区', '学院', '运营部', '小组', '员工岗位', '入职时间', '在职天数','月内工作天数','本月业绩','之前月份业绩'

        # 总业绩 / 当月日均业绩 / 总日均业绩
        col_names = ['总业绩', '本月日均业绩', '总日均业绩']
        result[col_names[0]] = result[cols[-1]] + result[cols[-2]]
        result[col_names[1]] = result[cols[-2]].divide(result[cols[-3]], fill_value=0).replace([np.nan, np.inf], 0)
        result[col_names[2]] = result[col_names[0]].divide(result[cols[-4]], fill_value=0).replace([np.nan, np.inf], 0)
        cols.extend(col_names)
        # '工号', '员工姓名', '地区', '学院', '运营部', '小组', '员工岗位', '入职时间', '在职天数','月内工作天数',
        # '本月业绩','之前月份业绩','总业绩','当月日均业绩','总日均业绩'

        # 是否新人 30天[包含] 是新人
        col_name3 = '是否新人'
        result[col_name3] = result[cols[-7]].map(lambda x: '是' if x < 31 else '')
        cols.append(col_name3)
        # '工号', '员工姓名', '地区', '学院', '运营部', '小组', '员工岗位', '入职时间', '在职天数','月内工作天数',
        # '本月业绩','之前月份业绩','总业绩','当月日均业绩','总日均业绩','是否新人'

        result.sort_values([cols[-3]], ascending=False, inplace=True)
        # 排名
        result.insert(0, '排名', list(range(1, result.shape[0] + 1)))
        result = FormatSeriesOfDf(df=result).format_dot(cols_name=[cols[-2], cols[-3]])
        # result[cols[-2]] = result[cols[-2]].map(lambda x: format(x, '.0f'))
        # result[cols[-3]] = result[cols[-3]].map(lambda x: format(x, '.0f'))
        result[cols[7]] = result[cols[7]].map(lambda x: x.strftime('%Y/%m/%d'))
        return result, self.start_date

    # 日报2 -- 1-推广小组组长当月与上月日均业绩对比 2-推广组组长当月业绩排名
    def groupleader_vs_lastmonth_avg(self):
        '''
         人员信息:当前数据库中人员表的人员信息
         '地区', '学院', '运营部', '小组', '员工姓名', '工号','员工岗位', '状态', '入职时间', '离职日期','接量类型', '在职天数'
         '''
        cols_ind = [5, 4, 0, 1, 2, 3, 8]
        cols = [self.inf_people_columns[_] for _ in cols_ind]  # '工号', '员工姓名', '地区', '学院', '运营部', '小组', '入职时间'
        # 小组组长信息
        inf = deepcopy(self.inf_groupleader[cols])
        inf[cols[-1]] = pd.to_datetime(inf[cols[-1]])

        # 当月工作天数
        col_names0 = ['本月工作天数', '上月工作天数']
        inf[col_names0[0]] = inf[cols[-1]].map(
            lambda x: MySqlServer(dq=self.dq, start_date=self.start_date,
                                  final_date=self.final_date).get_sum_holidays() if x < self.start_date_date else (
                MySqlServer(dq=self.dq, start_date=x,
                            final_date=self.final_date).get_sum_holidays() if self.start_date_date <= x <= self.final_date_date
                else 0))
        last_1m_1d_date = self.cls_final_date.get_date_Nmonthes_firstday(n=-1)
        last_1m_1d_str = get_str_date(last_1m_1d_date)
        last_1m_endd_date = self.cls_final_date.get_date_Nmonthes_endday(n=-1)
        last_1m_endd_str = get_str_date(last_1m_endd_date)
        cls_sqlserver_1 = MySqlServer(dq=self.dq, start_date=last_1m_1d_str, final_date=last_1m_endd_str)  # 上月数据类
        inf[col_names0[1]] = inf[cols[-1]].map(
            lambda x: cls_sqlserver_1.get_sum_holidays() if x < last_1m_1d_date else (
                MySqlServer(dq=self.dq, start_date=x,
                            final_date=last_1m_endd_str).get_sum_holidays() if last_1m_1d_date <= x <= last_1m_endd_date
                else 0))

        cols.extend(col_names0)  # '工号', '员工姓名', '地区', '学院', '运营部', '小组','入职时间','当月内工作天数', '上月内工作天数'

        # 本月业绩
        col_name3 = '本月业绩'
        tg_this_month = self.tg_peoples_groupby_0
        tg_this_month.rename(columns={'业绩': col_name3}, inplace=True)
        result = pd.merge(inf, tg_this_month, on=cols[0], how='left')
        cols.append(col_name3)  # '工号', '员工姓名', '地区', '学院', '运营部', '小组','入职时间','当月内工作天数', '上月内工作天数','本月业绩'

        # 上月业绩
        col_name2 = '上月业绩'
        tg_other_month = cls_sqlserver_1.get_tg_peoples_groupby()  # 上月员工业绩  工号 业绩
        tg_other_month.rename(columns={'业绩': col_name2}, inplace=True)
        result = pd.merge(result, tg_other_month, on=cols[0], how='left')
        cols.append(col_name2)  # '工号', '员工姓名', '地区', '学院', '运营部', '小组','入职时间','当月内工作天数', '上月内工作天数','本月业绩','上月业绩'

        # 日均量
        col_names1 = ['上月日均业绩', '本月日均业绩']
        result[col_names1[0]] = result[cols[-1]].divide(result[cols[-3]], fill_value=0).replace([np.nan, np.inf], 0)
        result[col_names1[1]] = result[cols[-2]].divide(result[cols[-4]], fill_value=0).replace([np.nan, np.inf], 0)
        cols.extend(col_names1)
        # '工号', '员工姓名', '地区', '学院', '运营部', '小组','入职时间','当月内工作天数', '上月内工作天数','本月业绩','上月业绩','上月日均业绩', '当月日均业绩'

        # 是否新人 30天[包含] 是新人
        col_name3 = '差值'
        result[col_name3] = result[cols[-1]] - result[cols[-2]]
        cols.append(col_name3)
        # '工号', '员工姓名', '地区', '学院', '运营部', '小组','入职时间','当月内工作天数', '上月内工作天数','本月业绩','上月业绩','上月日均业绩', '当月日均业绩','差值'

        # 1-推广小组组长当月与上月日均业绩对比
        new_cols1 = cols[:6] + cols[-3:]
        result1 = result[new_cols1].copy()
        result1.sort_values([cols[-1]], ascending=False, inplace=True)
        # 2-推广组组长当月业绩排名
        new_cols2 = cols[:6] + cols[7:8] + cols[9:10] + cols[-2:-1]
        result2 = result[new_cols2].copy()
        result2.sort_values([cols[-2]], ascending=False, inplace=True)

        # 序号
        col_name4 = '序号'
        result_all = [result1, result2]
        for result_one in result_all:
            result_one.insert(0, col_name4, list(range(1, result_one.shape[0] + 1)))
            # 格式
            _new_cols = list(result_one.columns)
            for i in range(-3, 0):
                result_one[_new_cols[i]] = result_one[_new_cols[i]].map(lambda x: format(x, '.0f'))

        # 均量 均线
        result_all_3 = int(result[cols[-2]].mean())
        result_all.extend([result_all_3, self.start_date])
        return result_all

    # 日报 -日阶段完成任务次数
    def days_finish(self):
        cols = self.inf_group_columns[:4]
        inf = self.inf_group[cols]  # '地区', '学院', '运营部','小组'

        col_name1 = '日期'
        col_name2 = '日'
        date_index = list(pd.date_range(self.start_date, self.final_date))
        date_index = pd.DataFrame(date_index, columns=[col_name1])
        new_index = cartesian_product_df(inf, date_index)  # 获取笛卡尔 索引 '地区', '学院', '运营部','小组' '日期'
        new_index[col_name2] = pd.to_datetime(new_index[col_name1]).dt.day
        cols.extend([col_name1, col_name2])  # '地区', '学院', '运营部','小组' '日期' '日'

        data_day_g = self.cls_sqlserver_0.get_tg_groups_groupby_date()
        data_day_g[col_name1] = pd.to_datetime(data_day_g[col_name1])
        result = pd.merge(left=new_index, right=data_day_g, how='left',
                          on=cols[:-1])  # '地区', '学院', '运营部','小组' '日期' '日' '创量'
        result[result.columns[-1]] = result[result.columns[-1]].fillna(0)
        del result[col_name1]
        cols = list(result.columns)  # '地区', '学院', '运营部','小组' '日' '创量'

        # 转置
        result.set_index(cols[:-1], inplace=True)
        result = result.unstack()  # 转置
        result.reset_index(inplace=True)
        result.columns = [col[1] if col[1] else col[0] for col in result.columns]  # 改变cols
        cols = list(result.columns)  # '地区', '学院', '运营部','1' '2'
        # 新增列
        col_names = [f'{_}达标天数' for _ in self.day_standard] + [f'{self.day_standard[0]}达标率', '应出勤天数']
        for i in range(2):
            result[col_names[i]] = result[cols[4:]].apply(lambda x: sum(x >= self.day_standard[i]), axis=1)
        result[col_names[-1]] = self.sum_workday
        result.insert(loc=(result.shape[1] - 1), column=col_names[-2], value=
        result[col_names[-3]].divide(result[col_names[-1]], fill_value=0).replace([np.nan, np.inf], 0))
        result = FormatSeriesOfDf(result).format_percentage(cols_name=col_names[-2])
        return result, self.start_date

    # 日报 -- 推广小组日均排名 详细到人员
    def groups_avg_rank_peoples(self):
        '''
           人员信息:当前数据库中人员表的人员信息
           '地区', '学院', '运营部', '小组', '员工姓名', '工号','员工岗位', '状态', '入职时间', '离职日期','接量类型', '在职天数'
           0        1       2       3       4           5       6           7       8       9           10          11
           '''
        cols = self.inf_people_columns
        # 在职员工信息
        inf_p = self.inf_people

        # 获取组内在职或本月离职员工
        cond0 = inf_p[cols[3]].notnull()
        cond2 = inf_p[cols[3]].str.endswith('组')
        cond3 = ((inf_p[cols[7]] == '在职') | (inf_p[cols[9]] >= self.start_date_date))
        inf_p = inf_p[cond0 & cond2 & cond3]
        # 删除状态
        # inf_p.drop(cols[-2], axis=1, inplace=True)

        # 月内工作天数
        col_name0 = '月内工作天数'
        inf_p1 = inf_p[inf_p[cols[7]] == '在职']
        inf_p1[col_name0] = inf_p1[cols[8]].map(
            lambda x: self.sum_workday if x < self.start_date_date else MySqlServer(dq=self.dq,
                                                                                    start_date=x,
                                                                                    final_date=self.final_date).get_sum_holidays())
        if sum(inf_p[cols[7]] == '离职'):
            inf_p2 = inf_p[inf_p[cols[7]] == '离职']
            inf_p2[col_name0] = inf_p2.apply(
                lambda x: MySqlServer(dq=self.dq,
                                      start_date=self.start_date,
                                      final_date=x[cols[9]]).get_sum_holidays() if x[cols[
                    8]] < self.start_date_date else MySqlServer(
                    dq=self.dq,
                    start_date=x[cols[8]],
                    final_date=self.final_date).get_sum_holidays())
            inf_p = pd.concat([inf_p2, inf_p1])
        else:
            inf_p = inf_p1
        cols.append(
            col_name0)  # '地区', '学院', '运营部', '小组', '员工姓名', '工号','员工岗位', '状态', '入职时间', '离职日期','接量类型', '在职天数','月内工作天数'

        # 得到本月业绩
        col_name3 = '本月业绩'
        tg_this_month = self.tg_peoples_groupby_0
        tg_this_month.rename(columns={'业绩': col_name3}, inplace=True)
        result = pd.merge(inf_p, tg_this_month, on=cols[5], how='left')
        result[col_name3].replace(np.nan, 0, inplace=True)
        cols.append(
            col_name3)  # '地区', '学院', '运营部', '小组', '员工姓名', '工号','员工岗位', '状态', '入职时间', '离职日期','接量类型', '在职天数','月内工作天数','本月业绩'

        # 得到小组日均量
        col_name4 = '小组日均量'
        result_groupby = fun_groupby(result, cols[:4], {cols[-1]: 'sum', cols[-2]: 'sum'})
        result_groupby[col_name4] = result_groupby[cols[-1]].divide(result_groupby[cols[-2]], fill_value=0).replace(
            [np.nan, np.inf], 0)
        # 得到排名
        col_name5 = '小组排名'
        result_groupby.sort_values(by=col_name4, ascending=False, inplace=True)
        result_groupby[col_name5] = list(range(1, result_groupby.shape[0] + 1))
        result_groupby = FormatSeriesOfDf(result_groupby).format_dot(cols_name=col_name4)
        result_groupby.drop(cols[-2:], axis=1, inplace=True)

        # 合并
        result = pd.merge(result, result_groupby, how='left', on=cols[:4])
        cols.extend([col_name4, col_name5])
        # '地区', '学院', '运营部', '小组', '员工姓名', '工号','员工岗位', '状态', '入职时间', '离职日期','接量类型', '在职天数','月内工作天数','本月业绩','小组日均量','小组排名'
        # 删除接量类型，状态，调整排名 列序
        result.drop([cols[7], cols[10]], axis=1, inplace=True)
        result.insert(0, col_name5, result.pop(col_name5))
        # 排序
        result.sort_values([col_name5, cols[6]], inplace=True)
        result[cols[8]], result[cols[9]] = [np.where(result[_].notnull(),
                                                     result[_].dt.strftime('%Y/%m/%d'),
                                                     '') for _ in cols[8:10]]
        return result, self.start_date

    # 晨晚报2 --  小组 和 运营部 两个结果
    def morning_evening(self, start_hour, end_hour):
        '''
        :param start_hour:
        :param end_hour:
        :return: 小组 和 运营部 两个结果
        '''
        if self.final_date_date == datetime.date.today():
            table_name = 'HourTg'
        else:
            table_name = 'Tg'
        cls_sqlserver = MySqlServer(dq=self.dq, start_date=self.final_date, final_date=self.final_date,
                                    table_name=table_name)
        # 业绩
        data = cls_sqlserver.get_tg_groups_groupby_date_hours(start_hour=start_hour,
                                                              end_hour=end_hour)
        cols_name = list(data.columns)  # '地区', '学院', '运营部', '小组', '日期','提交小时', '业绩'
        data_g = fun_groupby(df=data, groupby_cols=cols_name[:4],
                             cols_fun_dict={cols_name[-1]: 'sum'})  # '地区', '学院', '运营部', '小组', '业绩'
        data_t = fun_groupby(df=data_g, groupby_cols=cols_name[:3],
                             cols_fun_dict={cols_name[-1]: 'sum'})  # '地区', '学院', '运营部', '业绩'
        datas = [data_g, data_t]
        # 组织信息
        inf_g = self.inf_group.iloc[:, :4]
        inf_t = self.inf_team.iloc[:, :3]
        infs = [inf_g, inf_t]
        # 在职人数
        p_num = cls_sqlserver.get_people_num()  # 地区 学院 运营部 小组 日期 人数
        cols_name_pnum = list(p_num.columns)  # 地区 学院 运营部 小组 日期 人数
        p_num_g = p_num.drop(cols_name_pnum[-2], axis=1)  # 地区 学院 运营部 小组 人数
        p_num_t = fun_groupby(p_num_g, cols_name_pnum[:3], {cols_name_pnum[-1]: 'sum'})  # 地区 学院 运营部 人数
        p_nums = [p_num_g, p_num_t]

        reindexs = [[0,1,2,3,5,7,4,6],[0,1,2,4,5,3,6]]

        result = []
        cols_num = [4, 3]
        for i in range(2):
            o_inf = infs[i]
            o_p_num = p_nums[i]
            o_data = datas[i]
            o_cols_num = cols_num[i]
            reind = reindexs[i]

            o_result = merge_many_dfs(cols_name[:o_cols_num], [o_inf, o_p_num, o_data], how='left')
            o_result[cols_name[-1]].fillna(0, inplace=True)
            cols_name1 = ['人均推量', '差值']
            o_result[cols_name1[0]] = o_result[cols_name[-1]].divide(o_result[cols_name_pnum[-1]]).replace(
                [np.nan, np.inf], 0)
            o_result = FormatSeriesOfDf(o_result).format_dot(cols_name=cols_name1[0])
            o_result.sort_values(cols_name[-1], ascending=False, inplace=True)
            next_row=o_result[cols_name[-1]].shift(1).fillna(method='bfill',limit=1)
            o_result[cols_name1[1]] = o_result[cols_name[-1]]-next_row
            o_result=reindex_cols(o_result,reind)
            result.append(o_result)
        return result

if __name__ == '__main__':
    # df = ReportDataAsDf('保定', '2020/6/8').people_performance()
    df = ReportDataAsDf('保定', '2020/6/8').morning_evening(0, 24)
    # df = Component('燕郊', '2020/6/2').component_df()
    print(df)
