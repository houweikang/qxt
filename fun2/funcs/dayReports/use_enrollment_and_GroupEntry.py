#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/26 16:46:27
# @Author  : HouWk
# @Site    : 
# @File    : use_enrollment_and_GroupEntry.py
# @Software: PyCharm


from fun_erm_and_groentry_rate import er_and_ge_rate_data


def er_and_GE_R_report(dq,T_or_G=1):
    # group_data, team_data, region_data, colege_data = er_and_ge_rate_data(dq)
    df_list = list(er_and_ge_rate_data(dq))
    if T_or_G==1:
        T_df_list = df_list[:-1]
        for df in T_df_list:



