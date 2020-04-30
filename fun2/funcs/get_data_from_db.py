#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/29 13:59:26
# @Author  : HouWk
# @Site    : 
# @File    : get_data_from_db.py
# @Software: PyCharm
from fun_db_QXT import operate_db


def get_component_data(start_date, final_date, dq):
    '''
    :param start_date: 起始日期
    :param final_date: 终止日期
    :param dq: 地区
    :return: df[日期 工号]
    '''
    sql = '''SELECT cast([提交时间] as date) 日期,[课程顾问-员工号] 工号
        FROM [QXT].[dbo].[Tg] 
        where cast([提交时间] as date) between '{}' and '{}' and [课程顾问-所属地区] = '{}'
        '''.format(start_date, final_date, dq)
    return operate_db(sql)


def get_er_and_ge_rate_data(dq):
    '''
    :param dq: 地区
    :return:推广统计数据
    '''
    sql = '''SELECT [量类型],[所属学院] as 学院 ,[所属部门] as 地区 ,[所属战队] as 运营部,[所属分组] as 小组 
            ,sum(cast([数据量] as int)) as 业绩 ,sum(cast([进群量] as int)) as 进群量 
            ,sum(cast([注册量] as int)) as 注册量 
            FROM [QXT].[dbo].[推广统计] 
            where [所在岗位] like '推广专员%' and [所属部门] like '{}%' 
            group by [量类型],[所属学院],[所属部门],[所属战队],[所属分组] 
            having sum(cast([数据量] as int)) <> 0  and sum(cast([进群量] as int)) <> 0  
            and sum(cast([注册量] as int)) <> 0'''.format(dq)
    return operate_db(sql)

def get_groups_data(start_date, final_date, dq):
    '''
    :param start_date:
    :param final_date:
    :param dq:
    :return: 学院 地区 运营部 小组 日期 创量
    '''
    sql = '''SELECT [推广专员-所属学院] 学院,[推广专员-所属地区] 地区
              ,[推广专员-所属战队] 运营部,[推广专员-所属小组] 小组 
              ,cast([提交时间] as date) 日期,count(*) 创量 
          FROM [QXT].[dbo].[Tg] 
          where cast([提交时间] as date) between '{}' and '{}' 
                and [推广专员-所属小组] like '%组' 
                and [推广专员-所属地区] = '{}' 
          group by [推广专员-所属学院],[推广专员-所属地区]
              ,[推广专员-所属战队],[推广专员-所属小组],cast([提交时间] as date)
        '''.format(start_date, final_date, dq)
    return operate_db(sql)

def get_holiday(dq, start_date, final_date):
    '''
    dt  dq
    :param dq:
    :param start_date:
    :param final_date:
    :return:dt  dq
    '''
    sql = '''exec proc_holiday '%s','%s','%s' ''' % (dq, start_date, final_date)
    return operate_db(sql)