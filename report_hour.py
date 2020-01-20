import xlwings as xw
from my_datetime import MyDateTime
from my_database import DbQxt
from my_box import MyBox

class hour_report:
      def __init__(self):
            dts = '2019/11/23-10'  # 暂定
            self.db = DbQxt()
            # get  compare and today‘s time
            # if not MyDateTime().dt_date.weekday():
            #       dts = '%s-%s' % ( MyDateTime().str_date, MyDateTime().str_hour)
            # else:
            #       dts = '%s-%s' % ( MyDateTime().str_date, MyDateTime().str_hour)
            dts = MyBox().my_inputbox("小时报：当前日期-时刻", dts)
            define_dts = dts.split('-')
            self.current_dt = define_dts[0]
            self.current_time = int(define_dts[1])
            self.col_name=MyDateTime(define_dts[0]).str_date_monthday

      def select(self,dq,Task_volume):
            # 判断数据库中是否包含current_dt的推量
            check_sql = '''SELECT 1 FROM `tg_dates_group` where `日期` ='%s' limit 1''' % self.current_dt
            if self.db.select(check_sql)[0]:  # 如果存在，则不用导入当前表内数据，直接生成小时报
                  col_name = MyDateTime(self.last_dt).str_date_monthday
                  select_sql1= '''SELECT
                                `推广专员-所属学院` 学院,
                                `推广专员-所属地区` 地区,
                                `推广专员-所属战队` 战队,
                                `推广专员-所属小组` 小组,
                                sum( CASE WHEN cast( `提交时间` AS date ) = '%s' THEN 1 ELSE 0 END ) AS '%s',
                                sum( CASE WHEN cast( `提交时间` AS date ) = '%s' THEN 1 ELSE 0 END ) AS '今日'
                            FROM
                                tg_hours_people
                            WHERE
                                tg_hours_people.`推广专员-所属小组` LIKE '%s'
                                AND cast( tg_hours_people.`提交时间` AS date ) BETWEEN '%s' AND '%s'
                                AND cast( DATE_FORMAT( tg_hours_people.`提交时间`, '%s' ) AS signed ) < %d
                            UNION
                            SELECT
                                people_group.`所属学院` 学院,
                                people_group.`地区`,
                                people_group.`战队`,
                                people_group.`小组`,
                                0 AS '%s',
                                0 AS '今日'
                            FROM
                                people_group
                            WHERE
                                people_group.`小组` LIKE '%s'
                                AND people_group.`日期` = '%s' ''' \
                         % (self.last_dt,self.last_dt,self.current_dt,'%组',self.last_dt,self.current_dt
                            ,'%k',self.current_time,self.last_dt,'%组',self.last_dt)
                  select_sql='''SELECT 学院,地区,战队,ifnull(a.小组,concat('总计-',a.战队)) as '小组'
                                  ,sum( `%s` ) `%s`
                                  ,sum( `今日` ) `今日`
                                  ,( sum( `今日` ) - sum( `%s` ) ) AS '差值'
                                  ,concat( sum( `今日` ) / %d * 100, '%s' ) AS '日指标完成比（%d)'
                              FROM
                                  ( %s ) AS a
                              where 地区 ='%s'
                              GROUP BY
                                  学院,地区,战队,小组
                              with rollup
                              having 战队 like '%s' '''  \
                        % (self.last_dt,self.col_name,self.last_dt,Task_volume,'%',Task_volume,select_sql1,dq,'%战队')

            datas_cols = self.db.select(select_sql)
            datas = datas_cols[0]
            cols = datas_cols[1]
            self.db.commit()
            self.db.close()
            wb = xw.books.active
            try:
                  sht = wb.sheets.add(dq)
            except:
                  sht = wb.sheets(dq)
            for i in range(0, len(cols)):
                  sht[0, i].value = cols[i][0]
            sht.range('A2').value = datas

      def
      #
      # def hour_report(dq,Task_volume):
      #       dts = '2019/11/22-2019/11/23-10'  # 暂定
      #       db = DbQxt()
      #       #get  compare and today‘s time
      #       # if not MyDateTime().dt_date.weekday():
      #       #       dts = '%s-%s-%s' % (MyDateTime().str_pastday(2), MyDateTime().str_date, MyDateTime().str_hour)
      #       # else:
      #       #       dts = '%s-%s-%s' % (MyDateTime().str_pastday(1), MyDateTime().str_date, MyDateTime().str_hour)
      #       dts = MyBox().my_inputbox("小时报：对比日期-当前日期-时刻", dts)
      #       define_dts = dts.split('-')
      #       last_dt = define_dts[0]
      #       current_dt = define_dts[1]
      #       current_time = int(define_dts[2])
      #       #判断数据库中是否包含current_dt的推量
      #       check_sql='''SELECT 1 FROM `tg_dates_group` where `日期` ='%s' limit 1''' % current_dt
      #       if db.select(check_sql)[0]: #如果存在，则不用导入当前表内数据，直接生成小时报
      #             col_name=MyDateTime(last_dt).str_date_monthday
      #             select_sql1= '''SELECT
      #                           `推广专员-所属学院` 学院,
      #                           `推广专员-所属地区` 地区,
      #                           `推广专员-所属战队` 战队,
      #                           `推广专员-所属小组` 小组,
      #                           sum( CASE WHEN cast( `提交时间` AS date ) = '%s' THEN 1 ELSE 0 END ) AS '%s',
      #                           sum( CASE WHEN cast( `提交时间` AS date ) = '%s' THEN 1 ELSE 0 END ) AS '今日'
      #                       FROM
      #                           tg_hours_people
      #                       WHERE
      #                           tg_hours_people.`推广专员-所属小组` LIKE '%s'
      #                           AND cast( tg_hours_people.`提交时间` AS date ) BETWEEN '%s' AND '%s'
      #                           AND cast( DATE_FORMAT( tg_hours_people.`提交时间`, '%s' ) AS signed ) < %d
      #                       UNION
      #                       SELECT
      #                           people_group.`所属学院` 学院,
      #                           people_group.`地区`,
      #                           people_group.`战队`,
      #                           people_group.`小组`,
      #                           0 AS '%s',
      #                           0 AS '今日'
      #                       FROM
      #                           people_group
      #                       WHERE
      #                           people_group.`小组` LIKE '%s'
      #                           AND people_group.`日期` = '%s' ''' \
      #                    % (last_dt,last_dt,current_dt,'%组',last_dt,current_dt,'%k',current_time,last_dt,'%组',last_dt)
      #             select_sql='''SELECT 学院,地区,战队,ifnull(a.小组,concat('总计-',a.战队)) as '小组'
      #                             ,sum( `%s` ) `%s`
      #                             ,sum( `今日` ) `今日`
      #                             ,( sum( `今日` ) - sum( `%s` ) ) AS '差值'
      #                             ,concat( sum( `今日` ) / %d * 100, '%s' ) AS '日指标完成比（%d)'
      #                         FROM
      #                             ( %s ) AS a
      #                         where 地区 ='%s'
      #                         GROUP BY
      #                             学院,地区,战队,小组
      #                         with rollup
      #                         having 战队 like '%s' '''  \
      #                   % (last_dt,col_name,last_dt,Task_volume,'%',Task_volume,select_sql1,dq,'%战队')
      #       datas_cols = db.select(select_sql)
      #       datas = datas_cols[0]
      #       cols = datas_cols[1]
      #       db.commit()
      #       db.close()
      #       wb=xw.books.active
      #       try:
      #             sht = wb.sheets.add(dq)
      #       except:
      #             sht = wb.sheets(dq)
      #
      #       for i in range(0, len(cols)):
      #             sht[0, i].value = cols[i][0]
      #       sht.range('A2').value = datas

if __name__ == '__main__':
      hour_report('保定',1000)
      hour_report('济南',750)