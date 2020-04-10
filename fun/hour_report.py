import numpy as np
import pandas as pd
from mybox import inputbox, only_ok
from datetime import datetime, timedelta
from data_sql import data_from_Tg, data_from_HourTg, groups_from_PeopleList
from config import target, dq_list, root_path
from makedirs import create_folder_date


def fun_hour_report(date_hour, dq, df_target=target()):
    # date_hour = '2020-04-08 14'
    now = datetime.now()
    fdt = datetime.strptime(date_hour, '%Y-%m-%d %H')
    sdt = fdt - timedelta(hours=fdt.hour)
    h = fdt.hour
    # 创建文件夹 并获取路径
    dt = date_hour.split()[0]
    path = create_folder_date(root_path, dt)
    # 从数据库中取数据
    if fdt.day == now.day:
        date_df = data_from_HourTg(sdt, fdt, dq)
    else:
        date_df = data_from_Tg(sdt, fdt, dq)
    # 从数据库中取组信息
    group_date = groups_from_PeopleList(dq)
    # 与量数据合并，为显示全部小组
    group_date_result = group_date.iloc[:, :4]
    group_date_result['提交时间'] = np.nan
    date_df = pd.concat([group_date_result, date_df])
    # 得到组长 和 config中的 目标
    group_headman = group_date.iloc[:, :5]
    group_headman = pd.merge(group_headman, df_target, on=['学院', '地区', '战队', '小组'])
    group_headman['目标'] = group_headman.目标.astype('int')
    group_headman.reset_index(drop=True, inplace=True)
    # 转换格式，为后续聚合做准备
    date_df['提交时间'] = pd.to_datetime(date_df['提交时间'], format='%Y-%m-%d %H:%M:%S')
    date_df['hour'] = date_df['提交时间'].dt.hour
    # 得到总计和每个小时量
    date_df.loc[date_df['hour'] < h, '总计'] = 1
    date_df.loc[date_df['hour'] < 10, str(10)] = 1
    cols = [str(10)]
    if h > 10:
        for i in range(11, h + 1):
            date_df.loc[date_df['hour'] == (i - 1), str(i)] = 1
            cols.append(str(i))
    # 聚合得到分组后的累加量
    values = list(date_df.columns[-(h - 8):])
    index = list(date_df.columns[:4])
    groupby = date_df.groupby(index)
    groupby = groupby[values].sum()
    # 列合并 组长和目标
    groupby = pd.merge(groupby, group_headman, on=['学院', '地区', '战队', '小组'])
    # 差值
    groupby['差值'] = groupby.总计 - groupby.目标
    # 分类汇总 战队、地区、学院 量
    groupby_team = groupby.groupby(by=['学院', '地区', '战队']).sum()
    groupby_team.reset_index(inplace=True)
    groupby_team['小组'] = groupby_team['战队'] + '总计'
    groupby_dq = groupby.groupby(by=['学院', '地区']).sum()
    groupby_dq.reset_index(inplace=True)
    groupby_dq['战队'] = groupby_dq['地区'] + '总计'
    groupby_colege = groupby.groupby(by=['学院']).sum()
    groupby_colege.reset_index(inplace=True)
    groupby_colege['地区'] = groupby_colege['学院'] + '总计'
    result = pd.concat([groupby, groupby_team, groupby_dq, groupby_colege], join='outer', ignore_index=True, sort=True)
    # 完成率
    result['完成率'] = result.总计 / result.目标
    # 填充空值
    result.fillna('-', inplace=True)
    # 排序列名
    columns = ['学院', '地区', '战队', '小组', '组长', '总计', '目标', '完成率', '差值'] + cols
    result = result.reindex(columns=columns)
    # 改数值格式
    result.iloc[:, 8:] = result.iloc[:, 8:].astype('int32', copy=True)
    result.set_index(['学院', '地区', '战队', '小组'], inplace=True)
    # 写入excel
    file_fullname = '{}{}{}点小时报.xlsx'.format(path, dq, h)
    sheet_name = '{}{}点小时报'.format(dq, h)
    result.to_excel(file_fullname, sheet_name=sheet_name)


def hour_report():
    now = datetime.now()
    now_str = now.strftime('%Y-%m-%d %H')
    date_hour = inputbox('请输入要统计的日期和时刻', now_str)
    for dq in dq_list:
        fun_hour_report(date_hour, dq)
    only_ok()


if __name__ == '__main__':
    hour_report()
