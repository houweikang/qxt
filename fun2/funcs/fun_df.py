#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/5/8 9:38:33
# @Author  : HouWk
# @Site    : 
# @File    : fun_df.py
# @Software: PyCharm
from functools import reduce
import pandas as pd


def cartesian_product_df(df1, df2):
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


def change_cols_format(df, cols, function):
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


def reindex_cols(df, new_cols):
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


def format_series_percentage(series, n=0):
    fun = lambda x: format(x, '.' + str(n) + '%')
    return series.map(fun)


class FormatSeriesOfDf():
    def __init__(self, df):
        self.df = df

    def df_format(self, func, cols_ind=None, cols_name=None):
        if cols_ind:
            columns = self.df.columns
            if isinstance(cols_ind, (tuple, list)):
                for ind in cols_ind:
                    if not isinstance(ind, int):
                        raise Exception('cols_ind\'para type Must int')
                    column = columns[ind]
                    self.df[column] = self.df[column].map(func)
            elif isinstance(cols_ind, int):
                column = columns[cols_ind]
                self.df[column] = self.df[column].map(func)
            else:
                raise Exception('cols_ind\'para type Must int')
        if cols_name:
            if isinstance(cols_name, (tuple, list)):
                for col_name in cols_name:
                    self.df[col_name] = self.df[col_name].map(func)
            else:
                self.df[cols_name] = self.df[cols_name].map(func)
        return self.df

    def format_dict(self, func_cols_name_d=None, func_cols_ind_d=None):
        if func_cols_name_d:
            for func, col_name in func_cols_name_d:
                self.df_format(func=func, cols_name=col_name)
        if func_cols_ind_d:
            for func, col_ind in func_cols_ind_d:
                self.df_format(func=func, cols_ind=col_ind)

    def format_percentage(self, n=0, cols_ind=None, cols_name=None):
        func = eval(f'''lambda x: '%.{n}f%%' % (x * 100)''')
        return self.df_format(func=func, cols_ind=cols_ind, cols_name=cols_name)

    def format_dot(self, n=0, cols_ind=None, cols_name=None):
        func = eval(f'''lambda x: '%.{n}f' % x''')
        return self.df_format(func=func, cols_ind=cols_ind, cols_name=cols_name)


def format_df_percentage(df, cols_ind=None, cols_name=None, n=0):
    fun = lambda x: format(x, '.' + str(n) + '%')
    if cols_ind:
        columns = df.columns
        if isinstance(cols_ind, (tuple, list)):
            for ind in cols_ind:
                if not isinstance(ind, int):
                    raise Exception('cols_ind  type int')
                column = columns[ind]
                df[column] = df[column].map(fun)
        elif isinstance(cols_ind, int):
            column = columns[cols_ind]
            df[column] = df[column].map(fun)
        else:
            raise Exception('cols_ind\'para type Must int')
    elif cols_name:
        if isinstance(cols_name, (tuple, list)):
            for col_name in cols_name:
                df[col_name] = df[col_name].map(fun)
        else:
            df[cols_name] = df[cols_name].map(fun)
    else:
        raise Exception('cols_ind or cols_name Must One!')
    return df


def fun_groupby(df, groupby_cols, cols_fun_dict, rename_cols_dict=None):
    df = df.groupby(groupby_cols)
    df = df.agg(cols_fun_dict)
    if rename_cols_dict:
        df.rename(columns=rename_cols_dict, inplace=True)
    # df = df.to_frame()
    df.reset_index(inplace=True)
    return df


def fill_na(df, fill_ind, para_ind_list, tag='总计-'):
    for _ in para_ind_list:
        columns = list(df.columns)
        cond1 = df[columns[fill_ind]].isna()
        cond2 = df[columns[_]].notna()
        df[columns[fill_ind]][cond1 & cond2] = tag + df[columns[_]][cond1 & cond2]
    return df
