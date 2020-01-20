import xlwings as xw
from my_database import DbQxt
from my_datetime import MyDateTime
from my_tkinter import MyBox
import os



def insert_cols():
    cols = [['地区','战队','小组','日期']]
    wb = xw.books.active
    dt=wb.name[:8]  #工作簿名称前八位为人员名单日期
    dt=MyDateTime(dt)
    if dt.is_date:
        dt=dt.str_date
    else:
        MyBox().my_onlyok("工作簿名称不符合日期标准，请检测后再追加！")
        exit()
    sht = xw.sheets.active
    rng  = sht.range('A1').expand('table').value
    ind = rng[0].index('所属部门')
    if '地区' in rng[0]:
        col = rng[0].index('地区')
    else:
        col = len(rng[0])
    for rlist in rng[1:]:
        dep = rlist[ind].replace(' ','')
        dep = dep.split("=>")
        new_cols = [''] * 4
        for _ in dep:
            if _.endswith('战队'):
                new_cols[1] = _
            elif _.endswith('组'):
                new_cols[2] = _
            else:
                new_cols[0] = _
        new_cols[3] = dt
        cols.append(new_cols)
    # print(cols)
    sht[0,col].options(expand='table').value = cols
    db = DbQxt()
    # check date data exists
    ind_peopleid=rng[0].index('员工工号')
    peopleid = rng[1][ind_peopleid]
    select_sql = "select * from peoplelist where cast(日期 as date) ='%s' and 员工工号 ='%s' limit 1" % (dt, peopleid)
    # print(select_sql)
    if db.select(select_sql)[0]:
        MyBox().my_onlyok(dt + " 数据已存在！请检测后再追加！")
        exit()
    # insert into hourtg
    rng = xw.Range('A1').expand('table')
    list_rng = rng.value
    row = rng.rows.count
    col = rng.columns.count
    for i in range(1, row):
        list_rngOne=[]
        for _ in list_rng[i]:
            strcell = str(_)
            if strcell.endswith('.0'):
                list_rngOne.append(strcell[:-2])
            else:
                list_rngOne.append(strcell)
        insert_sql = "insert into peoplelist values(" + "'%s'," * (col - 1) + "'%s')"
        insert_sql = insert_sql % tuple(list_rngOne)
        # print(insert_sql)
        db.operate(insert_sql)
        update_sql="update peoplelist set 地区='%s' where 地区='%s'" % ('济南','济南推广一部')
        db.operate(update_sql)
    db.commit()
    db.close()
    print('OK')

if __name__ == '__main__':
    path=r'c:\Users\Administrator\Desktop\111'
    excel_lists=[]
    for dirpath,dirname,filenames in os.walk(path):
        for filename in filenames:
            excel_lists.append(os.path.join(dirpath,filename))

    app = xw.App(visible=True, add_book=False)
    app.screen_updating = True

    for excel_fullname in excel_lists:
        if excel_fullname.endswith('xls'):
            wb=app.books.open(excel_fullname)
            insert_cols()
            wb.save()
            wb.close()

    app.quit()