from db_QXT import operate_db


def groups_data_from_Tg(start_date, final_date, dq='%'):
    sql = '''SELECT [推广专员-所属学院] '学院', [推广专员-所属地区] '地区'
            , [推广专员-所属战队] '战队', [推广专员-所属小组] '小组', COUNT(*) '推广量'
            FROM dbo.Tg 
            WHERE ([推广专员-所属小组] LIKE '%组') and CONVERT(date, 提交时间) between '{}' and '{}' 
            and [推广专员-所属地区] = '{}'
            GROUP BY [推广专员-所属学院], [推广专员-所属地区], [推广专员-所属战队]
            , [推广专员-所属小组]'''.format(start_date, final_date, dq)
    # 返回 学院 地区 战队 小组 推广量
    return operate_db(sql)


def groups_data_from_HourTg(start_date, final_date, dq='%'):
    sql = '''SELECT [推广专员-所属学院] AS 学院, [推广专员-所属地区] AS 地区
            , [推广专员-所属战队] AS 战队, [推广专员-所属小组] AS 小组, COUNT(*) AS 推广量 
            FROM dbo.HourTg 
            WHERE ([推广专员-所属小组] LIKE '%组') and CONVERT(date, 提交时间) between '{}' and '{}' 
            GROUP BY [推广专员-所属学院], [推广专员-所属地区], [推广专员-所属战队]
            , [推广专员-所属小组]'''.format(start_date, final_date, dq)
    # 返回 学院 地区 战队 小组 推广量
    return operate_db(sql)


def groups_from_PeopleList(dq='%'):
    sql = '''SELECT a.所属学院 as 学院, a.地区, a.战队, a.小组, b.员工姓名 AS 组长, a.组内人数 
            FROM (SELECT 所属学院, 地区, 战队, 小组, COUNT(小组) AS '组内人数'
                 FROM  dbo.PeopleList 
                 WHERE (小组 LIKE '%组') AND (状态 = '在职') AND (地区 = '{}') 
                 GROUP BY 所属学院, 地区, 战队, 小组) AS a 
            LEFT OUTER JOIN 
                (SELECT distinct 所属学院, 地区, 战队, 小组, 员工岗位, 员工姓名 
                 FROM  dbo.PeopleList 
                 WHERE (员工岗位 = '推广专员组长') AND (状态 = '在职') AND (地区 = '{}')) AS b 
            ON a.所属学院 = b.所属学院 AND a.地区 = b.地区 AND a.战队 = b.战队  
                AND a.小组 = b.小组 '''.format(dq, dq)
    # 返回 学院 地区 战队 小组 组长 组内人数【在职】
    # 注意是否存在两个不同的在职组长
    return operate_db(sql)


def data_from_Tg(start_date, final_date, dq='%'):
    sql = '''SELECT [推广专员-所属学院] '学院', [推广专员-所属地区] '地区'
            , [推广专员-所属战队] '战队', [推广专员-所属小组] '小组', [提交时间] 
            FROM dbo.Tg 
            WHERE ([推广专员-所属小组] LIKE '%组') and CONVERT(date, 提交时间) between '{}' and '{}' 
            and [推广专员-所属地区] = '{}' '''.format(start_date, final_date, dq)
    return operate_db(sql)


def data_from_HourTg(start_date, final_date, dq='%'):
    sql = '''SELECT [推广专员-所属学院] '学院', [推广专员-所属地区] '地区'
            , [推广专员-所属战队] '战队', [推广专员-所属小组] '小组', [提交时间] 
            FROM dbo.HourTg 
            WHERE ([推广专员-所属小组] LIKE '%组') and CONVERT(date, 提交时间) between '{}' and '{}' 
            and [推广专员-所属地区] = '{}' '''.format(start_date, final_date, dq)
    return operate_db(sql)
