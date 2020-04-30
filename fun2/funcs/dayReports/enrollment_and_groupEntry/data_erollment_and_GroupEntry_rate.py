#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 15:18:15
# @Author  : HouWk
# @Site    : 
# @File    : data_erollment_and_GroupEntry_rate.py
# @Software: PyCharm
from config import dict_substituet
from get_data_from_db import get_er_and_ge_rate_data


def er_and_ge_rate_data(dq):
    # 获取推广统计数据
    data = get_er_and_ge_rate_data(dq)

    # 替换地区 部门 推广一部-'' 1战队-推广一部
    data['地区'].replace(r'推广\w部', '', regex=True, inplace=True)
    data['运营部'].replace(dict_substituet, inplace=True)

    deparment_list = ['量类型', '学院', '地区', '运营部', '小组']
    group_data = df_operation(data, deparment_list)
    team_data = df_operation(data, deparment_list[:-1])
    region_data = df_operation(data, deparment_list[:-2])
    colege_data = df_operation(data, deparment_list[:-3])

    return group_data, team_data, region_data, colege_data


def data_format(df):
    '''
    :param df: 业绩 注册量 进群量 及进群率 和 注册率格式化
    :return:
    '''
    df['业绩'] = df['业绩'].map(lambda x: '%.0f' % x)
    df['注册量'] = df['注册量'].map(lambda x: '%.0f' % x)
    df['进群量'] = df['进群量'].map(lambda x: '%.0f' % x)

    df['注册率'] = df['注册率'].map(lambda x: '%.0f%s' % (x * 100, '%'))
    df['进群率'] = df['进群率'].map(lambda x: '%.0f%s' % (x * 100, '%'))
    return df


def df_inset(df):
    '''
    :param df: 业绩 注册量 进群量 及进群率 和 注册率格式化
    :return:
    '''
    df['进群率'] = df['进群量'] / df['业绩']
    df['注册率'] = df['注册量'] / df['业绩']
    return df


def df_operation(data, cols):
    final_data = data.groupby(cols)
    final_data = final_data.sum()
    final_data.reset_index(inplace=True)
    final_data = df_inset(final_data)
    final_data = data_format(final_data)  # 调整格式
    new_cols = cols + ['业绩', '注册量', '注册率', '进群量', '进群率']
    final_data = final_data.reindex(new_cols, axis=1)
    return final_data


if __name__ == '__main__':
    er_and_ge_rate_data('济南')
