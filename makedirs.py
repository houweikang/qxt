import os
from datetime import datetime
import dateutil.parser

def create_folder(path):
    path=path.strip()
    path=path.rstrip('\\')
    if not os.path.exists(path):
        os.makedirs(path)

def create_folder_date(path,date):
    try:
        date=dateutil.parser.parse(date)
        date_fmt=datetime.strftime(date,'%Y%m%d')
        path_year=date_fmt[:4]
        path_month=date_fmt[:6]
        path_day=date_fmt
        path='''%s/%s/%s/%s/''' % (path,path_year,path_month,path_day)
        create_folder(path)

    except ValueError:
        print('未创建路径！')

def main():
    create_folder_date(r'e:\报表\晨报小时报模板\use\日报','2020/2/1')

if __name__ == '__main__':
    main()