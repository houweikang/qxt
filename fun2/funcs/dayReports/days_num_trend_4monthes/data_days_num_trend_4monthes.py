#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/27 13:36:34
# @Author  : HouWk
# @Site    : 
# @File    : data_days_num_trend_4monthes.py
# @Software: PyCharm

# !/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 15:18:15
# @Author  : HouWk
# @Site    :
# @File    : data_erollment_and_GroupEntry_rate.py
# @Software: PyCharm
import numpy as np
import pandas as pd

from config import dict_substituet
from fun_date import get_str_lastNmonth_firstday
from get_data_from_db import get_groups_data, get_holiday


def days_num_trend_4monthes(start_date, final_date, dq):
    data = get_groups_data(start_date, final_date, dq)

    dep = ['学院', '地区', '运营部', '小组']
    #学院
    dep_colege = dep[0:1]
    data_colege_list = get_group(data, dep_colege, dq, start_date, final_date)

    # 地区
    dep_region = dep[0:2]
    data_region_list = get_group(data, dep_region, dq, start_date, final_date)

    #运营部
    dep_team = dep[0:3]
    data_team_list = get_group(data, dep_team, dq, start_date, final_date)

    all_list = [data_colege_list, data_region_list, data_team_list]

    return all_list


def get_group(df, cols, dq, start_date, final_date):
    cols_index = df[cols].drop_duplicates()
    date = ['日期']
    date_index = list(pd.date_range(start_date, final_date))
    date_index = pd.DataFrame(date_index, columns=date)

    new_index = get_new_index(cols_index, date_index)  # 获取笛卡尔索引

    cols1 = cols + date
    df_groupby = df.groupby(cols1)
    df_sum = df_groupby['创量'].sum()

    df_sum = df_sum.reindex(new_index, fill_value=0)
    df_sum = df_sum.to_frame()
    df_sum.reset_index(inplace=True)
    df_sum['日期'] = pd.to_datetime(df_sum['日期'])

    # 获取工作日平均量
    current_month_fisrtday = get_str_lastNmonth_firstday(final_date)  # 获取本月1号
    df_holiday = get_holiday(dq, current_month_fisrtday, final_date)
    df_holiday['日期'] = pd.to_datetime(df_holiday['dt'])
    df_holiday = df_holiday[['日期', dq]]
    df_avg = pd.merge(df_sum, df_holiday, how='left', on='日期')
    df_avg = df_avg[(df_avg[dq] == 1) & (df_avg['创量'] != 0)]
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

    result_list = get_only_col(result, cols, df_avg)  # 唯一化输出到列表

    return result_list


def get_new_index(pd1, pd2):
    df1_col = list(pd1.columns)
    df2_col = list(pd2.columns)
    new_cols = df1_col + df2_col
    pd1['key'] = 1
    pd2['key'] = 1
    df = pd.merge(pd1, pd2, on='key')
    new_index = df[new_cols]
    return new_index


def get_only_col(df, cols, df_avg):
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
        avg = pd.merge(df_avg, avg_row, left_index=True, right_on=cols)['创量']
        avg_row.iloc[:, (len_cols + 1):] = avg
        new_df = pd.concat([new_df, avg_row])
        for colname in list(new_df.columns)[(len_cols + 1):]:
            new_df[colname] = new_df[colname].map(lambda x: '%.0f' % x if not isinstance(x, str) else '')
        if '运营部' in list(new_df.columns):
            new_df['运营部'] = new_df['运营部'].map(dict_substituet)

        result.append(new_df)
    return result


# if __name__ == '__main__':
#     days_num_trend_4monthes('2020/3/1', '2020/4/29', '保定')
