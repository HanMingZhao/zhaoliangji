import pymysql as db
import xlwt
import datetime
import configparser
import time
import calendar
import decimal

stime = time.time()

cf = configparser.ConfigParser()
cf.read('conf.conf')
dbhost = cf.get('db', 'db_host')
dbuser = cf.get('db', 'db_user')
dbport = cf.getint('db', 'db_port')
dbpass = cf.get('db', 'db_pass')
dbase = cf.get('db', 'db_db')
scon = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, charset='utf8')
scur = scon.cursor()

wb = xlwt.Workbook()
today = datetime.datetime.today()
yesterday = today-datetime.timedelta(1)
month = today.month
year = today.year
dateFormat = '%Y-%m-%d'
first = datetime.datetime.strptime(str(year)+'-'+str(month)+'-'+str(1), dateFormat)

print('日销量...', time.time()-stime)
sheet = wb.add_sheet('销售总计')
sheet.write(0, 0, '日期')
sheet.write(0, 1, '销量')
daySaleSql = '''
SELECT DATE(oo.`pay_at`),COUNT(1) FROM panda.`odi_order` oo
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`pay_at` > '{}'
AND oo.`pay_at` < '{}'
GROUP BY DATE(oo.`pay_at`)
'''

scur.execute(daySaleSql.format(first.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchall()
saleSum = 0
for i, r in enumerate(result):
    sheet.write(i+1, 0, r[0])
    sheet.write(i+1, 1, r[1])
    saleSum += r[1]
row = len(sheet.rows)
target = cf.getint('db', 'target')
sheet.write(row, 0, '总计')
sheet.write(row, 1, saleSum)
sheet.write(row, 2, '距离目标还差 {} 台'.format(target-saleSum))

path = cf.get('path', 'path')
wb.save(path+'day.xls')
scur.close()
scon.close()
print('overtime...', time.time()-stime)
