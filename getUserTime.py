import os
import psycopg2
import win32com.client
import datetime


def str_turn_date(string):
    return datetime.date(*map(int, str(string).split('-')))


def write_log(text):
    path = os.getcwd() + '\\log.txt'  # 获取当前工作路径的字符串
    # print(path)
    file = open(path, 'a')
    file.write(
        '\n--------\n' +
        '%s\n' % datetime.datetime.now() +
        '%s\n' % text
    )
    file.close()


def write_log_title(text):
    path = os.getcwd() + '\\log.txt'  # 获取当前工作路径的字符串
    file = open(path, 'a')
    file.write(text)
    file.close()


class TimeCard:

    def __init__(
            self,
            ip_list,
            start_datetime=None,
            end_datetime=None,
            database="cr",
            user="postgres",
            password="123456",
            host="127.0.0.1",
            port="5432"
    ):
        self.start_datetime = start_datetime
        self.end_datetime = end_datetime
        self.ip_list = ip_list
        self.conn = psycopg2.connect(database=database, user=user, password=password, host=host, port=port)
        self.sql = self.conn.cursor()

    # 获取开始时间与结束时间
    def get_date_time(self):
        start = None
        end = None
        # 获取前一天的日期
        if self.start_datetime is None and self.end_datetime is None:
            # print(start)
            # 如果没有传入指定的日期，则获取前一天的日期
            get_date = datetime.date.today()
            start = end = get_date + datetime.timedelta(-1)  # 获取前一天
        elif self.start_datetime is not None and self.end_datetime is None:
            # 如果之传入了开始日期，则获取开始日期
            start = end = str_turn_date(self.start_datetime)
        elif self.start_datetime is None and self.end_datetime is not None:
            # 如果之传入了结束日期，则获取结束日期
            start = end = str_turn_date(self.end_datetime)
        else:
            start = str_turn_date(self.start_datetime)
            end = str_turn_date(self.end_datetime)

        self.start_datetime = datetime.datetime.combine(start, datetime.time(00, 00, 00))
        self.end_datetime = datetime.datetime.combine(end, datetime.time(23, 59, 59))
        # print(self.start_datetime, self.end_datetime)

    def query_data(self, ip):
        sql_code = "SELECT * FROM kq where kq_date>='%s' and kq_date<='%s' and kq_ip='%s'" \
                   % (self.start_datetime, self.end_datetime, ip)
        self.sql.execute(sql_code)
        data = self.sql.fetchall()
        # print(data)
        return len(data) != 0

    def write_data(self, kq_name, kq_datetime, kq_date, kq_time, kq_ip):
        sql_code = "INSERT INTO kq (kq_name,kq_datetime,kq_date,kq_time,kq_ip)" \
                   " values ('%s','%s','%s','%s','%s')"\
                   % (kq_name, kq_datetime, kq_date, kq_time, kq_ip)
        self.sql.execute(sql_code)

    def get_user_time(self):
        print(self.start_datetime, self.end_datetime)
        zk = win32com.client.Dispatch('zkemkeeper.ZKEM.1')  # 获取中控API

        for ip in self.ip_list:

            if self.query_data(ip):
                text = '当前 %s 上的时间段 %s -- %s 的数据已有或已有部分，结束执行' % (ip, self.start_datetime, self.end_datetime)
                write_log(text)
                print(text)
                continue

            zk.Connect_Net(ip, 4370)

            # zk.ReadGeneralLogData(1)
            zk.ReadTimeGLogData(1, self.start_datetime, self.end_datetime)
            # zk.ReadTimeGLogData(1, '2022-7-5 00:00:00', '2022-7-5 23:59:59')

            while 1:
                # 获取打卡机数据
                exists, name_id, func, mode, year, month, day, hour, minute, second, work = zk.SSR_GetGeneralLogData(1)
                if not exists:
                    print('获取完毕')
                    break
                # print(exists, name_id, func, mode, year, month, day, hour, minute, second, work, ip)
                kq_datetime = datetime.datetime(year, month, day, hour, minute, second)
                kq_date = datetime.date(year, month, day)
                kq_time = datetime.time(hour, minute, second)
                data = {
                    'kq_name': name_id,
                    'kq_datetime': kq_datetime,
                    'kq_date': kq_date,
                    'kq_time': kq_time,
                    'kq_ip': ip
                }
                self.write_data(**data)
                print(name_id, kq_datetime, ip)

            zk.Disconnect()
            self.conn.commit()
            t = '当前 %s 上的时间段 %s -- %s 的数据以成功写入' % (ip, self.start_datetime, self.end_datetime)
            write_log(t)
            print(t)

        self.conn.close()

    def run(self):
        write_log_title('\n\n^^^^^^  %s 程序开始执行  ^^^^^^' % datetime.datetime.now())
        self.get_date_time()
        self.get_user_time()
        write_log_title('******  %s 程序执行结束  ******\n\n' % datetime.datetime.now())


if __name__ == '__main__':
    ip_lists = ['192.168.1.31', '192.168.1.32', '192.168.1.35', '192.168.1.36']
    # ip_lists = ['192.168.1.37']

    options = {
        'ip_list': ip_lists,
        # 'start_datetime': '2022-6-1',
        # 'end_datetime': '2022-7-6',
        'host': '192.168.1.90',
        'password': '310104'
    }

    run = TimeCard(**options)
    run.run()
