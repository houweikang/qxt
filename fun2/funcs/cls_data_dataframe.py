#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/12 10:05:34
# @Author  : HouWk
# @Site    : 
# @File    : cls_data_dataframe.py
# @Software: PyCharm
from functools import reduce

import numpy as np
import pyodbc
from pandas import DataFrame
import pandas as pd

from cls_date import MyDate, get_str_date


class DataAsDF:
    def __init__(self, dq, start_date, final_date):
        self.dq = dq
        self.final_date = final_date
        self.start_date = start_date
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

    def operate_DB(self, sql):  # 数据库
        dr = 'SQL Server Native Client 11.0'  # driver
        # sv = "192.168.1.43"  # 服务器名称
        sv = 'localhost'  # 服务器名称
        db = "QXT"  # '数据库名称
        un = "sa"  # '数据库连接用户名
        pw = "houweikang123"  # '数据库连接密码

        con_str = r'DRIVER={};SERVER={};DATABASE={};UID={};PWD={}'.format(
            dr, sv, db, un, pw)
        conn = pyodbc.connect(con_str)
        try:
            sql_data = pd.read_sql(sql, conn)
            return DataFrame(sql_data)
        except:
            cursor = conn.cursor()
            cursor.execute(sql)
            conn.commit()

    def fun_get_inf_from_proc(self, proc_name, *args):
        sql = '''exec %s %s ''' % (proc_name, args)
        sql = sql.replace('(', '').replace(')', '')
        if sql.endswith(', '):
            sql = sql[:-2]
        return self.operate_DB(sql)

    def get_teams_inf(self, proc_name='proc_inf_teams'):
        # 学院 地区 运营部 推广人数 战队长 战队长工号 组个数 小组每日目标 日目标 月目标
        return self.fun_get_inf_from_proc(proc_name, self.dq)

    def get_team_groupby_date(self,dq,start_date,final_date, proc_name='proc_sum_team'):
        #学院 地区 运营部 业绩
        return self.fun_get_inf_from_proc(proc_name, dq, start_date, final_date)

    def groups_groupby_data(self):
        '''

        :return: 地区 学院  运营部 小组 日期 创量
        '''
        sql = '''SELECT [推广专员-所属地区] 地区 ,[推广专员-所属学院] 学院,
                  [推广专员-所属战队] 运营部,[推广专员-所属小组] 小组 ,
                  cast([提交时间] as date) 日期,count(*) 创量 
              FROM [QXT].[dbo].[Tg] 
              where cast([提交时间] as date) between '{}' and '{}' 
                    and [推广专员-所属小组] like '%组' 
                    and [推广专员-所属地区] = '{}' 
              group by [推广专员-所属地区],[推广专员-所属学院],
                  [推广专员-所属战队],[推广专员-所属小组],cast([提交时间] as date)
            '''.format(self.start_date, self.final_date, self.dq)
        return self.operate_DB(sql)

    def get_holiday(self, dq, start_date, final_date):
        '''
        :return:dt dq
        '''
        sql = '''exec proc_holiday '%s','%s','%s' ''' % (dq, start_date, final_date)
        return self.operate_DB(sql)

    def get_sum_holidays(self, dq, start_date, final_date):
        sql = '''exec proc_sum_workdays '%s','%s','%s' ''' % (dq, start_date, final_date)
        return self.operate_DB(sql).iloc[0,0]

    def format_percentage(self,series, n=0):
        fun = lambda x: format(x, '.' + str(n) + '%')
        return series.map(fun)

    def reindex_cols(self,df, new_cols):
        '''
        保证长度一致
        :param df:
        :param new_cols: 列表 下标
        :return:
        '''
        cols_list = list(df.columns)
        if len(cols_list) == len(new_cols):
            new_cols_list = []
            for _ in new_cols:
                new_cols_list.append(cols_list[_])
            df = df.reindex(columns=new_cols_list)
            return df

    def merge_many_dfs(self,cols, dfs, how):
        return reduce(lambda left, right: pd.merge(left, right, on=cols, how=how), dfs)

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

    def erollment_and_group_entry_rate_1(self,data, cols):
        #erollment_and_group_entry_rate_df的辅助函数
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

    def erollment_and_group_entry_rate_df(self):
        # 获取推广统计数据
        sql = '''SELECT [量类型],[所属学院] as 学院 ,[所属部门] as 地区 ,[所属战队] as 运营部,[所属分组] as 小组 
                ,sum(cast([数据量] as int)) as 业绩 ,sum(cast([进群量] as int)) as 进群量 
                ,sum(cast([注册量] as int)) as 注册量 
                FROM [QXT].[dbo].[推广统计] 
                where [所在岗位] like '推广专员%' and [所属部门] like '{}%' 
                group by [量类型],[所属学院],[所属部门],[所属战队],[所属分组] 
                having sum(cast([数据量] as int)) <> 0  and sum(cast([进群量] as int)) <> 0  
                and sum(cast([注册量] as int)) <> 0'''.format(self.dq)
        data = self.operate_DB(sql)

        # 替换地区 部门 推广一部-'' 1战队-推广一部
        data['地区'].replace(r'推广\w部', '', regex=True, inplace=True)
        data['运营部'].replace(self.dict_substituet, inplace=True)

        deparment_list = ['量类型', '学院', '地区', '运营部', '小组']
        group_data = self.erollment_and_group_entry_rate_1(data, deparment_list)
        team_data = self.erollment_and_group_entry_rate_1(data, deparment_list[:-1])
        region_data = self.erollment_and_group_entry_rate_1(data, deparment_list[:-2])
        colege_data = self.erollment_and_group_entry_rate_1(data, deparment_list[:-3])
        return group_data, team_data, region_data, colege_data

    def trend_4monthes_df(self):
        # 获取小组业绩
        data = self.groups_groupby_data()

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

        all_list = [data_region_list,data_colege_list,  data_team_list]

        return all_list

    def trend_4monthes_1(self, df, cols):
        #trend_4monthes_df的辅助函数
        #获取列表集
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
        df_holiday['日期'] = pd.to_datetime(df_holiday['dt'])
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
        #trend_4monthes_df的辅助函数
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

    def vs_last_month_df(self):
        cols = ['学院', '地区', '运营部']
        inf_team = self.get_teams_inf()[cols]  # 获取运营部

        # 获取运营部本月业绩 和 均量
        data = self.vs_last_month_1(self.dq, self.start_date, self.final_date, inf_df=inf_team, cols=cols)

        # 获取上月业绩
        start_date_1 = get_str_date( MyDate(self.start_date).get_date_Nmonth_sameday(-1))
        final_date_1 = get_str_date(MyDate(self.final_date).get_date_Nmonth_sameday(-1))
        data_1 = self.vs_last_month_1(self.dq, start_date_1, final_date_1, inf_df=inf_team, cols=cols)

        result = pd.merge(data_1, data, how='left',on=cols)

        result['运营部'] = result['运营部'].map(self.dict_substituet)
        new_list_index=[0,1,2,5,9,6,10,3,4,7,8]
        result = self.reindex_cols(result,new_list_index)
        return result

    def vs_last_month_1(self, dq,start_date, final_date, inf_df, cols):
        data = self.get_team_groupby_date(dq,start_date, final_date)
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
        data = self.change_cols_format(df=data, cols=change_cols, function=func)
        return data

    def complete_rate_time_rate_df(self):
        cols = ['学院', '地区', '运营部', '月目标']
        # 获取运营部及月目标
        target_team = self.get_teams_inf()[cols]

        # 获取inf_cols
        inf_cols = cols[:-1]

        # 获取本月业绩
        data = self.get_team_groupby_date(self.dq,self.start_date, self.final_date)
        col = get_str_date(self.final_date, '%m') + '月'
        col_amount = col + '业绩'
        data = data.rename({'业绩': col_amount}, axis=1)

        # 获取完成率
        result = pd.merge(target_team, data, on=inf_cols, how='left')
        result['完成率'] = result[col_amount].divide(result[cols[-1]])
        result['完成率'] = self.format_percentage(result['完成率'], n=0)

        # 获取工作天数
        sum_workday = self.get_sum_holidays(self.dq,self.start_date, self.final_date)
        col_workday = col + '工作天数'
        result[col_workday] = sum_workday

        # 获取本月总的工作天数
        final_date_0 = get_str_date( MyDate(self.final_date).get_date_Nmonthes_endday())
        sum_workday_1 = self.get_sum_holidays(self.dq,self.start_date, final_date_0)
        col_workday_1 = col + '总工作天数'
        result[col_workday_1] = sum_workday_1

        # 获取时间消耗率
        col_complete_rate = '时间消耗率'
        result[col_complete_rate] = result[col_workday].divide(result[col_workday_1])
        result[col_complete_rate] = self.format_percentage(result[col_complete_rate], n=0)

        # 运营部
        result['运营部'] = result['运营部'].map(self.dict_substituet)

        # 改变列顺序
        new_list_index = [0, 1, 2, 5, 8, 3, 4, 6, 7]
        result = self.reindex_cols(result, new_list_index)
        return result

    def change_cols_format(self,df, cols, function):
        '''
        :param df: 原df
        :param cols: 改变的列 列表
        :param function: 改变的样式
        :return: 改变后的df
        '''
        for _ in cols:
            df[_] = df[_].map(function)
        return df


if __name__ == '__main__':
    df = DataAsDF('保定', '2020/5/1', '2020/5/10').complete_rate_time_rate_df()
    print(df)
