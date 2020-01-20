from datetime import datetime, timedelta
from dateutil.parser import parse


class MyDateTime:
    default_dt = datetime.strftime(datetime.now(), '%Y-%m-%d %H:00:00')
    def __init__(self, input_datetime=default_dt):
        self.str_inputdt = input_datetime
        self.dt_inputdt = parse(self.str_inputdt)
        self.week = self.dt_inputdt.weekday()
        self.month_days = self.dt_inputdt.day

    @property
    def is_date(self):
        """
        Return whether the string can be interpreted as a date.
        :param string: str, string to check for date
        :param fuzzy: bool, ignore unknown tokens in string if True
        """
        try:
            parse(self.str_inputdt)
            return True
        except ValueError:
            return False

    @property
    def dt_time(self):
        return self.dt_inputdt.time()

    @property
    def str_time(self):
        return self.dt_date.strftime('%H:%M:%S')

    @property
    def dt_hour(self):
        return self.dt_inputdt.hour

    @property
    def str_hour(self):
        return str(self.dt_hour)

    @property
    def dt_date(self):
        return self.dt_inputdt.date()

    @property
    def str_date(self):
        return self.dt_date.strftime('%Y/%m/%d')

    @property
    def str_date_monthday(self):
        return self.dt_date.strftime('%m月%d日')


    def dt_pastday(self, days=1):
        return self.dt_inputdt - timedelta(days=days)

    def str_pastday(self, days=1):
        return self.dt_pastday(days).strftime('%Y/%m/%d')

    @property
    def dt_monday(self):
        return self.dt_inputdt - timedelta(days=self.week)

    @property
    def str_monday(self):
        return self.dt_monday.strftime('%Y/%m/%d')

    @property
    def dt_currentmonth_date(self):
        return self.dt_inputdt - timedelta(days=(self.month_days-1))

    @property
    def str_currentmonth_date(self):
        return self.dt_currentmonth_date.strftime('%Y/%m/%d')


def main():
    # if not MyDateTime().dt_date.weekday():
    #     dts = '%s-%s-%s' % (MyDateTime().str_pastday(2), MyDateTime().str_date, MyDateTime().str_hour)
    # else:
    #     dts = '%s-%s-%s' % (MyDateTime().str_pastday(1), MyDateTime().str_date, MyDateTime().str_hour)

    # print(dts)
    print(MyDateTime('2020/1/1').is_date)


if __name__ == '__main__':
    main()