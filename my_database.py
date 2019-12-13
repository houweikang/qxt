import pymysql

class Database():
    """新建数据库类"""

    def __init__(self, host, port, database, charset, user, password):
        """连接数据库"""
        self.con = pymysql.connect(host=host, port=port,
                            database=database, charset=charset,
                            user = user, password = password)

    def select(self, sql):
        """筛选结果"""
        try:
            with self.con.cursor() as cursor:
                cursor.execute(sql)
                return cursor.fetchall(), cursor.description
        except:
            print('筛选有问题')

    def operate(self, sql):
        """delete or update"""
        try:
            with self.con.cursor() as cursor:
                result = cursor.execute(sql)
            # if result == 1:
            #     print('操作成功!')
        except:
            # db.rollback()
            print('有问题')

    def commit(self):
        self.con.commit()

    def close(self):
        self.con.close()

class DbQxt(Database):
    """继承Database 连接QXT数据库"""
    def __init__(self):
        host = 'localhost'
        port = 3306
        database = 'QXT'
        charset = 'utf8'
        user = 'root'
        password = 'houweikang123'
        Database.__init__(self,host, port, database, charset, user, password)


def main():
    db=DbQxt()
    sql='SELECT DISTINCT people_num.`小组` FROM people_num '
    print(db.select(sql))
    sql = 'Delete * from hourtg'
    db.operate(sql)
    db.close()


if __name__ == '__main__':
    main()


