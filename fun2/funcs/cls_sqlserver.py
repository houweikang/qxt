#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/20 9:05:01
# @Author  : HouWk
# @Site    : 
# @File    : cls_sqlserver.py
# @Software: PyCharm
from datetime import timedelta
from functools import reduce

import numpy as np
import pyodbc
from pandas import DataFrame
import pandas as pd
from cls_date import MyDate, get_str_date
from fun_df import change_cols_format, format_percentage, reindex_cols, fun_groupby, cartesian_product_df, fill_na


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
    def __init__(self, dq, start_date=None, final_date=None):
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
        return self.operate_DB(sql)

    def get_tg_data(self, table_name='Tg', dq=None, start_date=None, final_date=None):
        '''
        新增了 [提交日期] type(date) ; [提交小时] type(int)
        :param table_name: Tg or HourTg
        :return:
        '''
        if not (dq and start_date and final_date):
            dq, start_date, final_date = self.dq, self.start_date, self.final_date
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
                    and CAST([提交时间] as date) between '{}' and '{}' '''.format(table_name, dq, start_date,
                                                                              final_date)
        return self.operate_DB(sql)

    def get_tg_groups_groupby_date(self, dq=None, start_date=None, final_date=None):

        '''
        :return:'地区', '学院', '运营部', '小组', '日期', '创量'
        '''
        if not (dq and start_date and final_date):
            dq, start_date, final_date = self.dq, self.start_date, self.final_date
        cols = ['地区', '学院', '运营部', '小组', '提交日期']
        df = self.get_tg_data(dq=dq, start_date=start_date, final_date=final_date)[cols]
        cols = ['地区', '学院', '运营部', '小组', '日期']
        df.rename(columns={'提交日期': '日期'}, inplace=True)
        cond = df['小组'].str.endswith('组')
        df = df[cond]
        df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '创量'})
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

    def get_tg_teams_groupby(self, dq=None, start_date=None, final_date=None):
        # 地区 学院  运营部 业绩
        if not (dq and start_date and final_date):
            dq, start_date, final_date = self.dq, self.start_date, self.final_date

        cols = ['地区', '学院', '运营部', '小组']
        df = self.get_tg_data(dq=dq, start_date=start_date, final_date=final_date)[cols]
        cols = ['地区', '学院', '运营部']
        cond = df['小组'].str.endswith('组')
        df = df[cond]
        df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '业绩'})
        return df

    def get_tg_teams_groupby_day_evening(self, dq=None, start_date=None, final_date=None):
        # 地区 学院  运营部
        if not (dq and start_date and final_date):
            dq, start_date, final_date = self.dq, self.start_date, self.final_date

        df = self.get_tg_data(dq=dq, start_date=start_date, final_date=final_date)
        cond1 = ((df['小组'].str.endswith('组')) & (df['提交小时'] < 18))
        cond2 = ((df['小组'].str.endswith('组')) & (df['提交小时'] >= 18))
        cols = ['地区', '学院', '运营部', '提交小时']
        df_day = df[cond1][cols]
        df_evening = df[cond2][cols]
        result = [df_day, df_evening]
        result1 = []
        for df in result:
            cols = ['地区', '学院', '运营部']
            df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '业绩'})
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
        df_all = pd.concat([df_g,df_t,df_c,df_r])
        df_all = df_all.reindex(columns=list(df_g.columns))
        return df_all

    def get_tg_peoples_groupby(self, dq=None, start_date=None, final_date=None):
        # 工号 业绩
        if not (dq and start_date and final_date):
            dq, start_date, final_date = self.dq, self.start_date, self.final_date

        cols = ['工号']
        df = self.get_tg_data(dq=dq, start_date=start_date, final_date=final_date)[cols]
        cols = ['工号']
        df = fun_groupby(df, cols, {cols[-1]: 'count'}, {cols[-1]: '业绩'})
        return df

    def get_tgtj(self):
        '''
        推广统计
        :return:
        '''
        sql = '''SELECT [量类型],[所属学院] as 学院 ,[所属部门] as 地区 ,[所属战队] as 运营部,[所属分组] as 小组
                ,sum(cast([数据量] as int)) as 业绩 ,sum(cast([进群量] as int)) as 进群量
                ,sum(cast([注册量] as int)) as 注册量
                FROM [QXT].[dbo].[推广统计]
                where [所在岗位] like '推广专员%' and [所属部门] like '{}%'
                group by [量类型],[所属学院],[所属部门],[所属战队],[所属分组]
                having sum(cast([数据量] as int)) <> 0  and sum(cast([进群量] as int)) <> 0
                and sum(cast([注册量] as int)) <> 0'''.format(self.dq)
        return self.operate_DB(sql)

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
        return self.operate_DB(sql)

    def get_holiday(self, dq=None, start_date=None, final_date=None):
        if not (dq and start_date and final_date):
            dq = self.dq
            start_date = self.start_date
            final_date = self.final_date

        sql = '''select dt [日期],{} 
                from [dbo].[holidays] 
                where dt between '{}' and '{}'
        '''.format(dq, start_date, final_date)
        return self.operate_DB(sql)

    def get_sum_holidays(self, dq=None, start_date=None, final_date=None):
        if not (dq and start_date and final_date):
            dq = self.dq
            start_date = self.start_date
            final_date = self.final_date

        result = self.get_holiday(dq, start_date, final_date)
        result = result[dq].sum()
        return result

    def get_people_num(self, dq=None, start_date=None, final_date=None):
        if not (dq and start_date and final_date):
            dq = self.dq
            start_date = self.start_date
            final_date = self.final_date

        sql = f'''SELECT [地区]
                  ,[学院]
                  ,[战队] [运营部]
                  ,[小组]
                  ,[日期]
                  ,[人数]
          FROM [QXT].[dbo].[people_num]
          where [地区] = '{dq}' and [日期] between '{start_date}' and '{final_date}' '''
        return self.operate_DB(sql)

    def get_inf_group(self):
        '''
        :return: '地区', '学院', '运营部', '小组', '管理员', '工号', '推广人数', '小组日目标', '小组月目标'
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
        :return:
        '''
        # 其他信息
        df = self.get_inf_group()
        return self.get_inf_fun(df, 3, '推广专员战队长')

    def get_inf_colege(self):
        '''
        基于 运营部信息 得到 学院信息
        :return:
        '''
        # 其他信息
        df = self.get_inf_team()
        return self.get_inf_fun(df, 2, '学院管理员')

    def get_inf_region(self):
        '''
        基于 运营部信息 得到 学院信息
        :return:
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



class ReportDataAsDf(MySqlServer):
    def __init__(self, dq, start_date=None, final_date=None):
        super().__init__(dq=dq, start_date=start_date, final_date=final_date)
        self.dq = dq
        self.start_date = start_date
        self.final_date = final_date
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

    # 日报 -- 近4个月人均分量
    def component_df(self):
        sql = '''SELECT cast([提交时间] as date) 日期,[课程顾问-员工号] 工号
            FROM [QXT].[dbo].[Tg]
            where cast([提交时间] as date) between '{}' and '{}' and [课程顾问-所属地区] = '{}'
            '''.format(self.start_date, self.final_date, self.dq)
        data = self.operate_DB(sql)  # 日期 工号 课程顾问

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
        new_index = list(pd.date_range(self.start_date, self.final_date))
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
        return result

    def erollment_and_group_entry_rate_1(self, data, cols):
        # erollment_and_group_entry_rate_df的辅助函数
        final_data = data.groupby(cols)
        final_data = final_data.sum()
        final_data.reset_index(inplace=True)
        final_data['进群率'] = final_data['进群量'] / final_data['业绩']
        final_data['注册率'] = final_data['注册量'] / final_data['业绩']
        final_data['业绩'] = final_data['业绩'].map(lambda x: '%.0f' % x)
        final_data['注册量'] = final_data['注册量'].map(lambda x: '%.0f' % x)
        final_data['进群量'] = final_data['进群量'].map(lambda x: '%.0f' % x)
        final_data['注册率'] = final_data['注册率'].map(lambda x: '%.0f%s' % (x * 100, '%'))
        final_data['进群率'] = final_data['进群率'].map(lambda x: '%.0f%s' % (x * 100, '%'))
        new_cols = cols + ['业绩', '注册量', '注册率', '进群量', '进群率']
        final_data = final_data.reindex(new_cols, axis=1)
        return final_data

    # 日报 -- 注册率与进群率
    def erollment_and_group_entry_rate_df(self):
        # 获取推广统计数据
        data = self.get_tgtj()

        # 替换地区 部门 推广一部-'' 1战队-推广一部
        data['地区'].replace(r'推广\w部', '', regex=True, inplace=True)
        data['运营部'].replace(self.dict_substituet, inplace=True)

        # deparment_list = ['量类型', '学院', '地区', '运营部', '小组']
        deparment_list = ['量类型', '地区','学院',  '运营部', '小组']
        group_data = self.erollment_and_group_entry_rate_1(data, deparment_list)
        team_data = self.erollment_and_group_entry_rate_1(data, deparment_list[:4])
        # region_data = self.erollment_and_group_entry_rate_1(data, deparment_list[:-2])
        # colege_data = self.erollment_and_group_entry_rate_1(data, deparment_list[:-3])
        colege_data = self.erollment_and_group_entry_rate_1(data, deparment_list[:3])
        region_data = self.erollment_and_group_entry_rate_1(data, deparment_list[:2])
        return group_data, team_data, region_data, colege_data

    # 日报 -- 近4月每日业绩趋势
    def trend_4monthes_df(self):
        # 获取小组业绩
        data = self.get_tg_groups_groupby_date()

        dep = ['地区', '学院', '运营部', '小组']
        # 地区
        dep_colege = dep[0:1]
        data_region_list = self.trend_4monthes_1(data, dep_colege)

        # 学院
        dep_region = dep[0:2]
        data_colege_list = self.trend_4monthes_1(data, dep_region)

        # 运营部
        dep_team = dep[0:3]
        data_team_list = self.trend_4monthes_1(data, dep_team)

        all_list = [data_region_list, data_colege_list, data_team_list]

        return all_list

    def trend_4monthes_1(self, df, cols):
        # trend_4monthes_df的辅助函数
        # 获取列表集
        cols_index = df[cols].drop_duplicates()
        date = ['日期']
        date_index = list(pd.date_range(self.start_date, self.final_date))
        date_index = pd.DataFrame(date_index, columns=date)

        # 获取笛卡尔索引
        df1_col = list(cols_index.columns)
        df2_col = list(date_index.columns)
        new_cols = df1_col + df2_col
        cols_index['key'] = 1
        date_index['key'] = 1
        new = pd.merge(cols_index, date_index, on='key')
        new_index = new[new_cols]
        # 聚合
        cols1 = cols + date
        df_groupby = df.groupby(cols1)
        df_sum = df_groupby['创量'].sum()
        # 重新设置所以
        df_sum = df_sum.reindex(new_index, fill_value=0)
        df_sum = df_sum.to_frame()
        df_sum.reset_index(inplace=True)
        df_sum['日期'] = pd.to_datetime(df_sum['日期'])

        # 获取工作日平均量
        current_month_fisrtday = get_str_date(MyDate(self.final_date).get_date_Nmonthes_firstday(n=0))  # 获取本月1号
        df_holiday = self.get_holiday(self.dq, current_month_fisrtday, self.final_date)
        df_holiday['日期'] = pd.to_datetime(df_holiday['日期'])
        df_holiday = df_holiday[['日期', self.dq]]
        df_avg = pd.merge(df_sum, df_holiday, how='left', on='日期')
        df_avg = df_avg[(df_avg[self.dq] == 1) & (df_avg['创量'] != 0)]
        df_avg = df_avg.groupby(cols)
        df_avg = df_avg['创量'].mean()

        df_sum['月'] = df_sum.日期.dt.month
        df_sum['日'] = df_sum.日期.dt.day

        cols2 = cols + ['月', '日', '创量']
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
        result_list = self.trend_4monthes_2(result, cols, df_avg)  # 唯一化输出到列表
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
            if '运营部' in list(new_df.columns):
                new_df['运营部'] = new_df['运营部'].map(self.dict_substituet)
            result.append(new_df)
        return result

    # 日报-- 与上月最高业绩对比
    def vs_last_month_df(self):
        cols = ['学院', '地区', '运营部']
        inf_team = self.get_inf_team()[cols]  # 获取运营部

        # 获取运营部本月业绩 和 均量
        data = self.vs_last_month_1(self.dq, self.start_date, self.final_date, inf_df=inf_team, cols=cols)

        # 获取上月业绩
        start_date_1 = get_str_date(MyDate(self.start_date).get_date_Nmonth_sameday(-1))
        final_date_1 = get_str_date(MyDate(self.final_date).get_date_Nmonth_sameday(-1))
        data_1 = self.vs_last_month_1(self.dq, start_date_1, final_date_1, inf_df=inf_team, cols=cols)

        result = pd.merge(data_1, data, how='left', on=cols)

        result['运营部'] = result['运营部'].map(self.dict_substituet)
        new_list_index = [0, 1, 2, 5, 9, 6, 10, 3, 4, 7, 8]
        result = reindex_cols(result, new_list_index)
        return result

    def vs_last_month_1(self, dq, start_date, final_date, inf_df, cols):
        data = self.get_tg_teams_groupby(dq, start_date, final_date)
        col = get_str_date(final_date, '%m') + '月'

        col_amount = col + '业绩'
        data = data.rename({'业绩': col_amount}, axis=1)

        # 获取工作天数
        sum_workday = self.get_sum_holidays(dq, start_date, final_date)
        col_workday = col + '工作天数'
        data[col_workday] = sum_workday

        # 获取日均业绩
        col_day_avg = col + '日均业绩'
        data[col_day_avg] = data[col_amount].div(sum_workday)
        data[col_day_avg] = data[col_day_avg].replace(np.inf, 0)

        # 获取全部运营部
        data = pd.merge(inf_df, data, on=cols, how='left')

        # 获取日均业绩均线
        col_day_avg_line = col + '日均业绩均线'
        data[col_day_avg_line] = data[col_day_avg].mean()

        data = data.fillna(0)

        func = lambda x: '%.0f' % x
        change_cols = [col_amount, col_workday, col_day_avg, col_day_avg_line]
        data = change_cols_format(df=data, cols=change_cols, function=func)
        return data

    # 日报 -- 完成率与时间消耗率
    def complete_rate_time_rate_df(self):
        cols = ['地区', '学院', '运营部', '月目标']
        # 获取运营部及月目标
        target_team = self.get_inf_team()[cols]

        # 获取inf_cols
        inf_cols = cols[:-1]

        # 获取本月业绩
        data = self.get_tg_teams_groupby(self.dq, self.start_date, self.final_date)
        col = get_str_date(self.final_date, '%m') + '月'
        col_amount = col + '业绩'
        data = data.rename({'业绩': col_amount}, axis=1)

        # 获取完成率
        result = pd.merge(target_team, data, on=inf_cols, how='left')
        result['完成率'] = result[col_amount].divide(result[cols[-1]])
        result['完成率'] = format_percentage(result['完成率'], n=0)

        # 获取工作天数
        sum_workday = self.get_sum_holidays(self.dq, self.start_date, self.final_date)
        col_workday = col + '工作天数'
        result[col_workday] = sum_workday

        # 获取本月总的工作天数
        final_date_0 = get_str_date(MyDate(self.final_date).get_date_Nmonthes_endday())
        sum_workday_1 = self.get_sum_holidays(self.dq, self.start_date, final_date_0)
        col_workday_1 = col + '总工作天数'
        result[col_workday_1] = sum_workday_1

        # 获取时间消耗率
        col_complete_rate = '时间消耗率'
        result[col_complete_rate] = result[col_workday].divide(result[col_workday_1])
        result[col_complete_rate] = format_percentage(result[col_complete_rate], n=0)

        # 运营部
        result['运营部'] = result['运营部'].map(self.dict_substituet)

        # 改变列顺序
        new_list_index = [0, 1, 2, 5, 8, 3, 4, 6, 7]
        result = reindex_cols(result, new_list_index)
        return result

    # 日报 -- 组长业绩贡献率
    def group_leader_rate(self):
        cols = ['地区', '学院', '运营部', '小组', '管理员', '工号']
        inf_team = self.get_inf_group()[cols]  # 获取组长信息

        # 获取本月组内人员业绩
        groupby_people = self.get_tg_peoples_groupby()
        # 获取本月组长业绩
        result = pd.merge(inf_team, groupby_people, on=cols[-1], how='left')
        result = result.rename({'业绩': '组长业绩', '管理员': '组长'}, axis=1)
        # 获取本月小组业绩
        groups_cols = cols[:4]
        groups_df = self.get_tg_groups_groupby()
        result = pd.merge(result, groups_df, on=groups_cols, how='left')
        result = result.rename({'创量': '小组业绩'}, axis=1)
        result[['组长业绩', '小组业绩']] = result[['组长业绩', '小组业绩']].fillna(0)
        result['组长贡献率'] = result['组长业绩'].divide(result['小组业绩'], fill_value=0)
        result['组长贡献率'] = result['组长贡献率'].replace(np.inf, 0)
        result['组长贡献率'] = format_percentage(result['组长贡献率'], n=0)
        result['运营部'] = result['运营部'].map(self.dict_substituet)

        # 改变列顺序
        new_list_index = [0, 1, 2, 3, 6, 7, 8, 4, 5]
        result = reindex_cols(result, new_list_index)
        result.sort_values(by='小组业绩', ascending=False, inplace=True)
        return result

    # 日报-- 运营部白天夜间业绩对比
    def team_day_evening(self):
        cols = ['学院', '地区', '运营部']
        inf_team = self.get_inf_team()[cols]  # 获取运营部

        # 获取运营部本月业绩
        day_df, evening_df = self.get_tg_teams_groupby_day_evening()

        day_df.rename({'业绩': '白天业绩'}, axis=1, inplace=True)
        evening_df.rename({'业绩': '夜间业绩'}, axis=1, inplace=True)

        data = pd.merge(day_df, evening_df, how='outer', on=cols)

        data['总业绩'] = data['白天业绩'] + data['夜间业绩']
        avg_teams = data['总业绩'].mean()

        # 获取工作天数
        sum_workday = self.get_sum_holidays()
        data['工作天数'] = sum_workday
        # 获取日均业绩
        col_day_avg = '日均业绩'
        data['日均业绩'] = data['总业绩'].div(sum_workday)
        data['日均业绩'] = data['日均业绩'].replace(np.inf, 0)
        data['日均业绩'] = data['日均业绩'].map(lambda x: format(x, '.0f'))
        # 获取白日占比
        data['白日占比'] = data['白天业绩'].divide(data['总业绩'])
        data[col_day_avg] = data[col_day_avg].replace(np.inf, 0)

        # 获取全部运营部
        data = pd.merge(inf_team, data, on=cols, how='left')

        # 获取均线
        data['业绩均线'] = avg_teams
        data = data.fillna(0)

        data['白日占比'] = format_percentage(data['白日占比'])

        data['运营部'] = data['运营部'].map(self.dict_substituet)
        new_list_index = [0, 1, 2, 4, 3, 5, 7, 9, 8, 6]
        result = reindex_cols(data, new_list_index)
        return result

    # 日报-- 近6个月业绩对比
    def team_six_month(self):
        cols = ['地区', '学院', '运营部', '小组', '日期']
        # 组
        # 获取小组
        inf = self.get_inf_group()[cols[:-1]]
        # 获取全部索引
        date_index = list(pd.date_range(self.start_date, self.final_date))
        date_index = pd.DataFrame(date_index, columns=[cols[-1]])
        cartesian_df = cartesian_product_df(inf, date_index)
        # 筛选符合要求的索引
        cartesian_df['月'] = cartesian_df[cols[-1]].dt.month
        cartesian_df['日'] = cartesian_df[cols[-1]].dt.day
        current_day = int(get_str_date(self.final_date, '%d'))
        col_name = '列名'
        month_list = [MyDate(self.final_date).get_date_Nmonth_sameday(i) for i in range(-5, 1)]
        col_dict = {dt.month: f'{dt.month}月1日-{dt.month}月{dt.day}日' for dt in month_list}
        cartesian_df[col_name] = cartesian_df['月'].map(col_dict)
        cond1 = cartesian_df['日'] <= current_day
        cartesian_df = cartesian_df[cond1][cols + [col_name]]

        # 获取小组近6月同期业绩
        group_df = self.get_tg_groups_groupby_date(dq=self.dq, start_date=self.start_date, final_date=self.final_date)
        col_name1 = cols[-1]
        group_df[col_name1] = pd.to_datetime(group_df[col_name1])
        result_g = pd.merge(cartesian_df, group_df, on=cols, how='left')
        col_name2 = result_g.columns[-1]
        result_g[col_name2] = result_g[col_name2].fillna(0)
        result_g = fun_groupby(result_g, list(result_g.columns)[:-1], {col_name2: 'sum'})

        # 获取小组近6月同期除休假外的工作总人数
        people_num_df = self.get_people_num(dq=self.dq, start_date=self.start_date, final_date=self.final_date)
        holiday_df = self.get_holiday(dq=self.dq, start_date=self.start_date, final_date=self.final_date)
        people_num_df = pd.merge(people_num_df, holiday_df, on=[col_name1])
        cond2 = people_num_df[self.dq] == 1
        people_num_df = people_num_df[cond2][people_num_df.columns[:-1]]
        people_num_df[col_name1] = pd.to_datetime(people_num_df[col_name1])

        # 合并 业绩 和 日期段内在职总人数
        result_group = pd.merge(result_g, people_num_df, how='left', on=cols)
        col_name3 = '人数'
        result_group[col_name3] = result_group[col_name3].fillna(0)
        del result_group[col_name1]
        result_group = fun_groupby(result_group, list(result_group.columns)[:5], {'创量': 'sum', col_name3: 'sum'})
        result_group['运营部'] = result_group['运营部'].map(self.dict_substituet)

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
        return all_result

    def team_six_month_fun(self, result):
        result_columns0 = list(result.columns)
        result['日人均创量'] = result[result_columns0[-2]].divide(result[result_columns0[-1]], fill_value=0)
        result['日人均创量'] = result['日人均创量'].fillna(0)
        result['日人均创量'] = result['日人均创量'].map(lambda x: format(x, '.0f'))
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

    # 日报-- 运营部夜间业绩对比
    def team_evening(self):
        cols = ['学院', '地区', '运营部']
        inf_team = self.get_inf_team()[cols]  # 获取运营部
        # 获取运营部本月业绩
        day_df, evening_df = self.get_tg_teams_groupby_day_evening()
        evening_df.rename({'业绩': '夜间业绩'}, axis=1, inplace=True)
        avg_teams = evening_df['夜间业绩'].mean()
        # 获取全部运营部
        data = pd.merge(inf_team, evening_df, on=cols, how='left')
        # 获取均线
        data['业绩均线'] = avg_teams
        data = data.fillna(0)
        data['运营部'] = data['运营部'].map(self.dict_substituet)
        return data

    # 日报 -- 运营部人员信息
    def team_peoplelist(self):
        start_date = self.start_date
        final_date = self.final_date

        start_date_date = MyDate(start_date).date_date
        final_date_date = MyDate(final_date).date_date

        l1_start_date = MyDate(start_date).get_date_Nmonth_sameday(-1)
        l1_final_date = MyDate(final_date).get_date_Nmonth_sameday(-1)

        df = self.get_Peoplelist()
        cond1 = df['小组'].notnull()
        df = df[cond1]
        cond = df['小组'].str.endswith('组')
        cols = ['地区', '学院', '运营部', '小组', '状态', '入职时间', '离职日期']
        df = df[cond][cols]
        df['入职时间'] = pd.to_datetime(df.入职时间)
        df['离职日期'] = df['离职日期'].fillna(final_date)
        df['离职日期'] = pd.to_datetime(df.离职日期)
        df['入职时间'] = df.入职时间 + timedelta(days=-0.5)
        df['在职天数'] = (df['离职日期'] - df['入职时间'] + timedelta(days=1)).dt.days
        columns = cols[:4] + ['在职人数', '在职人数(在职天数>180)', '在职人数(在职天数 91-180)', '在职人数(在职天数 61-90)',
                              '在职人数(在职天数 31-60)', '在职人数(在职天数 7-30)', '在职人数(在职天数 1-6)',
                              '离职人数(在职天数>180)', '离职人数(在职天数 91-180)', '离职人数(在职天数 61-90)',
                              '离职人数(在职天数 31-60)', '离职人数(在职天数 7-30)', '离职人数(在职天数 1-6)',
                              f'入职人数（{final_date}）', '本月入职人数', '本月离职人数', '本月新进员工离职人数',
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
        result_group['运营部'] = result_group['运营部'].map(self.dict_substituet)
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
            r[columns[-3]] = format_percentage(r[columns[-7]].divide(r[columns[-9]]))  # 当月新人离职率
            r[columns[-2]] = format_percentage(r[columns[-8]].divide(r[columns[-8]] + r[columns[4]]))  # '本月离职率'
            r[columns[-1]] = format_percentage(r[columns[-5]].divide(r[columns[-5]] + r[columns[-4]]))  # '本月离职率'
            del r[columns[-5]]
            del r[columns[-4]]

        result_group, result_team, result_colege, result_region = r_ls
        # 合并
        result_all = pd.concat(r_ls, ignore_index=True)
        result_all = result_all.reindex(columns=list(result_group.columns), fill_value='')
        result_all.sort_values(by=columns[:4],inplace=True)
        #填充总计
        result_all=fill_na(result_all,3,[2,1,0])
        #填充不必要的单元格
        result_all[columns[:3]] = result_all[columns[:3]].fillna('-')
        return result_all

    # 日报 -- 目标完成率
    def complete_rate_df(self):
        cols = ['地区', '学院', '运营部','小组', '月目标']
        # 获取小组及以上架构和目标
        inf_all = self.get_inf_all()[cols]

        # 获取本月所有架构业绩
        data = self.get_tg_all_groupby()
        col_name3 = list(data.columns)[-1]
        col_name = '本月业绩'
        data.rename(columns={col_name3:col_name},inplace=True)
        # 合并
        result = pd.merge(left=inf_all,right=data,how='left',on=cols[:4])
        result[cols[2]] = result[cols[2]].map(self.dict_substituet)
        #排序
        result.sort_values(by=cols[:4], inplace=True)
        # 填充总计-
        result = fill_na(result,3,[2,1,0])
        result[cols[:3]] = result[cols[:3]].fillna('-')
        # 距目标差值
        col_name1 = '距目标差值'
        result[col_name1] = result[col_name] - result[cols[-1]]
        col_name2 = '完成率'
        result[col_name2] = format_percentage(result[col_name].divide(result[cols[-1]]))
        #小数点后0位
        columns = list(result.columns)
        for _ in [4,6]:
            result[columns[_]] = result[columns[_]].map(lambda x:format(x,'.0f'))
        return result

if __name__ == '__main__':
    df = ReportDataAsDf('保定', '2020/5/1', '2020/5/28').complete_rate_df()
    print(df)
