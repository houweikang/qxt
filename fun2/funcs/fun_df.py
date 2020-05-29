#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/8 9:38:33
# @Author  : HouWk
# @Site    : 
# @File    : fun_df.py
# @Software: PyCharm
from functools import reduce
import pandas as pd



def cartesian_product_df(df1,df2):
    '''
    返回笛卡尔乘积后的df
    :param df:
    :param df2:
    :return:
    '''
    df1 = df1.drop_duplicates()
    df2 = df2.drop_duplicates()
    df1_col = list(df1.columns)
    df2_col = list(df2.columns)
    new_cols = df1_col + df2_col
    df1['key'] = 1
    df2['key'] = 1
    new = pd.merge(df1, df2, on='key')
    return new[new_cols]

def change_cols_format(df,cols,function):
    '''
    :param df: 原df
    :param cols: 改变的列 列表
    :param function: 改变的样式
    :return: 改变后的df
    '''
    for _ in cols:
        df[_] = df[_].map(function)
    return df

def merge_many_dfs(cols, dfs, how):
    return reduce(lambda left, right: pd.merge(left, right, on=cols, how=how), dfs)

def reindex_cols(df,new_cols):
    '''
    保证长度一致
    :param df:
    :param new_cols: 列表 下标
    :return:
    '''
    cols_list = list(df.columns)
    if len(cols_list) == len(new_cols):
        new_cols_list =  []
        for _ in new_cols:
            new_cols_list.append(cols_list[_])
        df = df.reindex(columns=new_cols_list)
        return df

def format_percentage(series,n=0):
    fun = lambda x: format(x, '.' + str(n) + '%')
    return series.map(fun)

def fun_groupby( df, groupby_cols, cols_fun_dict,rename_cols_dict=None):
    df = df.groupby(groupby_cols)
    df = df.agg(cols_fun_dict)
    if rename_cols_dict:
        df.rename(columns=rename_cols_dict, inplace=True)
    # df = df.to_frame()
    df.reset_index(inplace=True)
    return df

def fill_na(df,fill_ind,para_ind_list,tag='总计-'):
    for _ in para_ind_list:
        columns = list(df.columns)
        cond1 = df[columns[fill_ind]].isna()
        cond2 = df[columns[_]].notna()
        df[columns[fill_ind]][cond1 & cond2] = tag + df[columns[_]][cond1 & cond2]
    return df