import xlwings as xw
from my_database import DbQxt
from my_datetime import MyDateTime
from my_tkinter import MyBox

def insert_cols():
    #today
    dt = MyDateTime().str_date
    dt = MyBox().my_inputbox("日期:", dt)
    # dt='2019/12/5'

    cols = [['地区','战队','小组','日期']]
    wb = xw.books.active
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
    select_sql = "select * from peoplelist where cast(日期 as date) ='%s' and 员工工号 ='%s'" % (dt, peopleid)
    if db.select(select_sql):
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

    db.commit()
    db.close()
    print('OK')
if __name__ == '__main__':
    insert_cols()
