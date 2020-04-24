#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/20 15:24:33
# @Author  : HouWk
# @Site    : 
# @File    : fun_component.py
# @Software: PyCharm
from db_QXT import operate_db
import pandas as pd


class ComponentData():
    def __init__(self, start_date, final_date,dq):
        '''
        返回 账号类型 学院 地区 战队 小组 推广专员 工号 日期 小时 推广量
        :param start_date:
        :param final_date:
        '''
        self.start_date = start_date
        self.final_date = final_date
        self.dq = dq
        sql = '''SELECT cast([提交时间] as date) 日期,[课程顾问-员工号] 工号
		    FROM [QXT].[dbo].[Tg] 
		    where cast([提交时间] as date) between '{}' and '{}' and [课程顾问-所属地区] = '{}'
            '''.format(self.start_date, self.final_date,self.dq)
        self.data = operate_db(sql)

    # def get_consultant_num(self):
        unique_consultant = self.data.drop_duplicates()
        self.consultant_num = unique_consultant.groupby('日期').count()
        self.consultant_num.rename(columns = {'工号':'顾问人数'},inplace=True)


    # def get_component_num(self):
        self.component_num = self.data.groupby('日期').count()
        self.component_num.rename(columns={'工号': '分量'},inplace=True)


    def get_avg_cs_cp(self):
        self.avg_cs_cp = pd.merge(self.consultant_num,self.component_num,on='日期')
        return self.avg_cs_cp


if __name__ == '__main__':
    ComD = ComponentData('2020/4/1','2020/4/19','燕郊')
    print(ComD.get_result())