#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 15:18:15
# @Author  : HouWk
# @Site    : 
# @File    : fun_erm_and_groentry_rate.py
# @Software: PyCharm
from db_QXT import operate_db
from use_substituet import dict_substituet


def er_and_ge_rate_data(dq):
    sql = '''SELECT [量类型],[所属学院] as 学院 ,[所属部门] as 地区 ,[所属战队] as 部门,[所属分组] as 小组 
            ,sum(cast([数据量] as int)) as 业绩 ,sum(cast([进群量] as int)) as 进群量 
            ,sum(cast([注册量] as int)) as 注册量 
            FROM [QXT].[dbo].[推广统计] 
            where [所在岗位] like '推广专员%' and [所属部门] like '{}%' 
            group by [量类型],[所属学院],[所属部门],[所属战队],[所属分组] 
            having sum(cast([数据量] as int)) <> 0  and sum(cast([进群量] as int)) <> 0  
            and sum(cast([注册量] as int)) <> 0'''.format(dq)
    data = operate_db(sql)

    # 替换地区 部门 推广一部-'' 1战队-推广一部
    data['地区'].replace(r'推广\w部', '', regex=True, inplace=True)
    data['部门'].replace(dict_substituet, inplace=True)

    deparment_list = ['量类型', '学院', '地区', '部门', '小组']
    group_data = df_operation(data, deparment_list)
    team_data = df_operation(data, deparment_list[:-1])
    region_data = df_operation(data, deparment_list[:-2])
    colege_data = df_operation(data, deparment_list[:-3])

    return group_data, team_data, region_data, colege_data
    # col_names = list(data.columns)
    # result_value = list(data.values)
    # return result_value, col_names


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