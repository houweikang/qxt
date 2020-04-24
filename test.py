import os
import pandas as pd


def walk(no_work_file_path,path):
    dic_nowork = pd.read_excel(no_work_file_path)
    dic_nowork = dic_nowork.ix[dic_nowork.状态 == '离职',[1,10,14]]
    dic_nowork['离职日期'] = pd.to_datetime(dic_nowork['离职日期'],format='%Y/%m/%d')
    p_datas = []
    if not os.path.exists(path):
        return -1
    for root, dirs, names in os.walk(path):
        for filename in names:
            filepath = os.path.join(root, filename)  # 路径和文件名连接构成完整路径
            dt = '{}-{}-{}'.format(filename[:4], filename[4:6], filename[6:8])
            p_data = pd.read_excel(filepath)
            p_data['日期'] = dt
            p_data['日期'] = pd.to_datetime(p_data['日期'],format='%Y-%m-%d')
            p_datas.append(p_data)
    result_all = pd.concat(p_datas,sort=False)[['接量类型','员工工号','员工姓名',
                                            '所属学院','所属部门','入职时间',
                                            '员工岗位','状态','日期']]
    cols = result_all.shape[1]
    result_dup = result_all.iloc[:,:(cols-1)]
    result_dup.sort_values(by='状态',inplace=True)
    result_dup.drop_duplicates(result_dup.columns[:-2],inplace=True)
    # result_dup.to_excel('bd202001.xlsx')

    result = pd.merge(result_dup,dic_nowork,how='left',on=['员工工号','状态'])

    date_min = result_all['日期'].dt.min()

    # result['开始日期'] =

    result.to_excel('bd202001.xlsx')


if __name__ == '__main__':
    no_work_file_path = r'e:\报表\所有数据信息\员工信息\保定员工信息\202004\20200420保定员工信息.xls'
    path = r'e:\报表\所有数据信息\员工信息\保定员工信息\2020\202001'
    walk(no_work_file_path,path)