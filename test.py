# import win32com.client
import datetime
#
# # ip_lists = ['192.168.1.31', '192.168.1.32', '192.168.1.35', '192.168.1.36', '192.168.1.37']
# ip_lists = ['192.168.1.37']
#
# zk = win32com.client.Dispatch('zkemkeeper.ZKEM.1')  # 获取中控API
#
# for ip in ip_lists:
#
#     zk.Connect_Net(ip, 4370)
#
#     # zk.ReadGeneralLogData(1)
#     # zk.ReadTimeGLogData(1, self.start_datetime, self.end_datetime)
#     # print(zk.ReadGeneralLogData(1))
#     print(zk.ReadTimeGLogData(1, '2022-7-5 00:00:00', '2022-7-5 23:59:59'))
#
#     while 1:
#         # 获取打卡机数据
#         exists, name_id, func, mode, year, month, day, hour, minute, second, work = zk.SSR_GetGeneralLogData(1)
#         if not exists:
#             print('获取完毕')
#             break
#         print(exists, name_id, func, mode, year, month, day, hour, minute, second, work, ip)
#
#     zk.Disconnect()


# import psycopg2
# date1 = datetime.datetime(2022, 1, 1, 0, 0, 0)
# date2 = datetime.datetime(2022, 1, 2, 23, 59, 59)
# print(date1 < date2)
# conn = psycopg2.connect(database="cr", user="postgres", password="123456", host="127.0.0.1", port="5432")


# 数据库插入语句
# cur = conn.cursor()

# cur.execute("SELECT * FROM kq where date>'2022-07-05 00:00:00' and date<'2022-07-05 23:59:59'")
# conn.execute("SELECT * FROM kq where date>='%s' and date<='%s' and ip='%s'" % (date1, date2, '192.168.1.36'))
# cur.execute("select * from kq where date>='2022-7-1 00:00:00' and date<='2022-7-5 23:59:59'")
# cur.execute("INSERT INTO kq (NAME,DATE,IP) values ('%s','%s','%s')" % ('5124', date2, '192.168.1.36'))

# s = cur.fetchall()
# print(len(s))
# print(s)

# print(1 <= 0)

# conn.commit()
# conn.close()

import os

# path = os.getcwd() + '\\log.txt'  # 获取当前工作路径的字符串
# # print(path)
# file = open(path, 'a')
# file.write(
#     'Hello world!\n' +
#     '-----'
# )
# file.close()

# print(datetime.datetime.now())

print(datetime.date(2022,11,2))
print(datetime.time(22,11,2))
