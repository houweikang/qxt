#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/16 9:45:34
# @Author  : HouWk
# @Site    : 
# @File    : per_capita.py
# @Software: PyCharm
from funcs.fun_date import get_str_lastNmonth_firstday, get_str_date, get_date_date
from funcs.fun_db_QXT import operate_db
from funcs.fun_chart import smooth_chart
import matplotlib.pyplot as plt
import numpy as np


def per_capita(str_final_date, dq):
    month = -3
    date_final_date = get_date_date(str_final_date)
    days = date_final_date.day
    str_final_date = get_str_date(str_final_date)
    str_l4m_firstday = get_str_lastNmonth_firstday(str_final_date, month)

    sql1 = "select [月] ,isnull([1],0) as '1日'"
    sql3 = "[1]"
    for i in range(2, 32):
        sql1 += ",isnull([{}],0) as '{}日'".format(i, i)
        sql3 += ",[{}]".format(i)

    sql = '''{} from ( 
    select (datename(MM,c.日期) + '月') AS 月,datename(d,c.日期) AS 日,(b.{}得到推量 / c.{}人数) As {}人均分量 
    from (select 日期,count(*) as {}人数 
    from (SELECT distinct cast([提交时间] as date) AS 日期,[课程顾问-员工号] 
    FROM [QXT].[dbo].[Tg] 
    where cast([提交时间] as date) between '{}' and '{}' 
        and [课程顾问-所属地区] = '{}') a 
    group by 日期) c, 
    (SELECT cast([提交时间] as date) 日期,count(*) as {}得到推量 
    FROM [QXT].[dbo].[Tg] 
    where cast([提交时间] as date) between '{}' and '{}' 
        and [课程顾问-所属地区] = '{}'  
    group by cast([提交时间] as date)) b 
    where c.日期 = b.日期 
    ) d PIVOT (sum({}人均分量) FOR [日] IN ({})) AS pvt order by [月]'''.format(sql1, dq, dq, dq, dq, str_l4m_firstday,
                                                                          str_final_date, dq,
                                                                          dq, str_l4m_firstday, str_final_date, dq, dq,
                                                                          sql3)
    sht_name = 'QX{}-近4个月每日人均分量趋势'.format(dq)
    table_title = sht_name
    sub_title = '{}-{}'.format(str_l4m_firstday, str_final_date)
    char_title = sht_name
    data = operate_db(sql)
    # rs = data.shape[0]
    cs = data.shape[1]
    clear_col = days - cs + 1
    # data.loc[-1,-17:] = None
    data.iloc[-1, clear_col:] = None
    x = np.arange(1,32)
    plt.rcParams['font.sans-serif'] = ['SimHei']  # SimHei黑体
    chart_x = data.columns[-31:]
    chart_y1 = data.iloc[1, -31:]
    smooth_chart(x, chart_y1)
    # plt.plot(x, chart_y1)
    # plt.show()


if __name__ == '__main__':
    per_capita('2020/4/15', '燕郊')
