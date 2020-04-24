#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/7 14:19:27
# @Author  : HouWk
# @Site    : 
# @File    : db_QXT.py
# @Software: PyCharm

import pyodbc
from pandas import DataFrame
import config
import pandas as pd


def conn_db():
    # dr  driver
    # sv  数据库服务器名称
    # db  数据库名称
    # un  数据库连接用户名
    # pw  数据库连接密码
    con_str = r'DRIVER={};SERVER={};DATABASE={};UID={};PWD={}'.format(
        config.dr, config.sv, config.db, config.un, config.pw)
    conn = pyodbc.connect(con_str)
    return conn


def operate_db(sql):
    conn = conn_db()
    sql_start = sql[:10].strip().lower()
    if sql_start.startswith('select'):
        sql_data = pd.read_sql(sql, conn)
        return DataFrame(sql_data)
    else:
        cursor = conn.cursor()
        cursor.execute(sql)
        conn.commit()

if __name__ == '__main__':
    sql = '''SELECT TOP 1000 [dt]
          ,[jn]
          ,[bd]
      FROM [QXT].[dbo].[holidays]'''
    print(operate_db(sql))