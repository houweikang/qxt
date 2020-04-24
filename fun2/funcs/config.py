#!/usr/bin/env python
# -*- coding= utf-8 -*-
# @Time    = 2020/4/7 10=34=19
# @Author  = HouWk
# @Site    = 
# @File    = config.py
# @Software= PyCharm

from pandas import DataFrame
import numpy as np
import pandas as pd

# 保存根目录
root_path = r"e:\python_Reports"

# 数据库
dr = 'SQL Server Native Client 11.0'  # driver
sv = "192.168.1.43"  # 数据库服务器名称
db = "QXT"  # '数据库名称
un = "sa"  # '数据库连接用户名
pw = "houweikang123"  # '数据库连接密码
#
# # 目标
# dq_list = ['保定', '济南']
#
#
# def target():
#     # 获取所有目标
#     msxy_bd_1team = msxy_bd_1team_group_target()
#     msxy_bd_2team = msxy_bd_2team_group_target()
#     msxy_jn_1team = msxy_jn_1team_group_target()
#     df_target = pd.concat([msxy_bd_1team, msxy_bd_2team, msxy_jn_1team])
#     return df_target
#
#
# # 美术学院 保定 1战队 小组目标
# def msxy_bd_1team_group_target():
#     team = ['美术学院', '保定', '1战队', ]
#     groups = ['1组', '2组', '3组', '4组', '5组', ]
#     target = 1000
#     result = df_target(team, groups, target)
#     return result
#
#
# # 美术学院 保定 2战队 小组目标
# def msxy_bd_2team_group_target():
#     team = ['美术学院', '保定', '2战队', ]
#     groups = ['1组', '2组', '3组', '4组', '5组', ]
#     target = 500
#     result = df_target(team, groups, target)
#     return result
#
#
# # 美术学院 济南 1战队 小组目标
# def msxy_jn_1team_group_target():
#     team = ['美术学院', '济南', '1战队', ]
#     groups = ['1组', '2组', '3组', '4组']
#     target = 750
#     result = df_target(team, groups, target)
#     return result
#
# # 生成 DataFrame
# def df_target(team, groups, target):
#     l_g = len(groups)
#     if isinstance(target, int):
#         target = [target] * l_g
#     l_tm = len(team)
#     tm = np.array(team * l_g).reshape(l_g, l_tm)
#     gp = np.array(groups).reshape(l_g, 1)
#     tg = np.array(target).reshape(l_g, 1)
#     result = np.hstack([tm, gp, tg])
#     result = DataFrame(result, columns=['学院', '地区', '战队', '小组', '目标'])
#     result.set_index(['学院', '地区', '战队', '小组'], inplace=True)
#     return result
