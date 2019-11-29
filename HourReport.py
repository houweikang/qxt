import xlwings as xw
from my_datetime import MyDateTime
from my_database import DbQxt
from my_tkinter import MyBox

# starttime = datetime.datetime.now()

def hour_report(dq):
      #delete from hourtg
      db = DbQxt()
      del_sql = 'Delete from temp_tg_hour'
      db.operate(del_sql)

      #insert into hourtg
      wb = xw.books.active
      rng = xw.Range('A1').expand('table')
      list_rng = rng.value
      row = rng.rows.count
      col = rng.columns.count

      for i in range(1, row):
            insert_sql = "insert into temp_tg_hour values(" + "'%s',"*(col-1) + "'%s')"
            insert_sql = insert_sql % tuple(list_rng[i])
            db.operate(insert_sql)

      db.commit()

      #get  compare and today‘s time
      # if not MyDateTime().dt_date.weekday():
      #       dts = '%s-%s-%s' % (MyDateTime().str_pastday(2), MyDateTime().str_date, MyDateTime().str_hour)
      # else:
      #       dts = '%s-%s-%s' % (MyDateTime().str_pastday(1), MyDateTime().str_date, MyDateTime().str_hour)

      dts = '2019/11/23-2019/11/25-10'

      dt = MyBox().my_inputbox("小时报：对比日期-当前日期-时刻", dts)
      define_dts = dt.split('-')
      last_dt = define_dts[0]
      current_dt = define_dts[1]
      current_time = int(define_dts[2]) - 1


      select_sql='''(SELECT a.`所属学院`,a.`地区`,a.`战队`,a.`小组`,a.`组长` 
            ,sum(case b.`日期` when '%s' then b.`所有推广量` else 0 end) as '上日' 
            ,sum(case b.`日期` when '%s' then b.`所有推广量` else 0  end) as '今日' 
            ,(sum(case b.`日期` when '%s' then b.`所有推广量` else 0  end)-sum(case b.`日期` when '%s' then b.`所有推广量` else 0 end)) as '差值' 
            FROM temp_group as a left join tg_hours as b 
            on a.`所属学院`=b.`推广专员-所属学院` and a.`地区`=b.`推广专员-所属地区` and a.`战队`=b.`推广专员-所属战队` and a.`小组`=b.`推广专员-所属小组` 
            where a.`地区`='%s' 
            and b.`时刻` between 0 and %d 
            and b.`日期` between '%s' and '%s'  
            group by a.`所属学院`,a.`地区`,a.`战队`,a.`小组`,a.`组长` 
            union 
            SELECT c.`所属学院`,c.`地区`,c.`战队`,concat('汇总-' ,c.`战队`) as '小组',c.`战队长` as 组长 
            ,sum(case d.`日期` when '%s' then d.`所有推广量` else 0 end) as '上日' 
            ,sum(case d.`日期` when '%s' then d.`所有推广量` else 0  end) as '今日' 
            ,(sum(case d.`日期` when '%s' then d.`所有推广量` else 0  end)-sum(case d.`日期` when '%s' then d.`所有推广量` else 0 end)) as '差值' 
            FROM temp_team as c 
            left join tg_hours as d 
            on c.`所属学院`=d.`推广专员-所属学院` and c.`地区`=d.`推广专员-所属地区` and c.`战队`=d.`推广专员-所属战队` 
            where c.`地区`='%s' and d.`时刻` between 0 and %d 
            and d.日期 between '%s' and '%s' 
            group by c.`所属学院`,c.`地区`,c.`战队`,c.`战队长`) 
            order by `所属学院`,`地区`,`战队`,`小组`''' \
            % (last_dt, current_dt, current_dt, last_dt, dq, current_time, last_dt, current_dt,
               last_dt, current_dt, current_dt, last_dt, dq, current_time, last_dt, current_dt)
      datas = db.select(select_sql)[0]
      cols = db.select(select_sql)[1]
      db.commit()
      db.close()
      sht = wb.sheets[dq]
      # sht = wb.sheets.add(name=dq)  # 新加sheet在开始
      for i in range(0, len(cols)):
            sht[0, i].value = cols[i][0]
      sht.range('A2').value = datas




hour_report('保定')
# hour_report('济南')


# endtime = datetime.datetime.now()
# print(endtime-starttime)