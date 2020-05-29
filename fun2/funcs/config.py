#!/usr/bin/env python
# -*- coding= utf-8 -*-
# @Time    = 2020/4/7 10=34=19
# @Author  = HouWk
# @Site    = 
# @File    = config.py
# @Software= PyCharm


# # 保存根目录
# root_path = r"e:\python_Reports"
#
# # 数据库
# dr = 'SQL Server Native Client 11.0'  # driver
# # sv = "192.168.1.43"  # 数据库服务器名称
# sv = 'localhost' # 数据库服务器名称
# db = "QXT"  # '数据库名称
# un = "sa"  # '数据库连接用户名
# pw = "houweikang123"  # '数据库连接密码

# 目标
day_reports_list = ['保定', '济南']
# component_dqs = ['济南', '燕郊']

# 近4月+平均线 颜色
RGBs = [(255, 255, 0), (0, 112, 192), (0, 176, 80), (255, 0, 0), (255, 255, 255)]  # 黄 蓝 绿 红 白

# 分量线颜色
component_RGB = RGBs[:4]

#白天夜间分线
hour_line = 18 #18点之前的业绩为白天业绩

#替换战队
# dict_substituet = {
#     '1战队': '运营一部',
#     '2战队': '运营二部',
#     '3战队': '运营三部',
#     '4战队': '运营四部',
#     '5战队': '运营五部',
#     '6战队': '运营六部',
#     '7战队': '运营七部',
#     '8战队': '运营八部',
# }

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
