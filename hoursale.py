import pymysql as db
import configparser
import datetime
import xlwt
import time
import numpy as np

start_time = time.time()
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

date_format = '%Y-%m-%d'
today = datetime.datetime.today()
days = [today]
for i in range(8):
    day = today - datetime.timedelta(i+1)
    days.insert(0, day)

sheet = wb.add_sheet('sheet1')
sheet.write(0, 0, '日期')
for i in range(24):
    sheet.write(0, i+1, i+1+'点')
for i, day in enumerate(days):
    query_sql = '''
    SELECT DATE(oo.`create_at`),HOUR(oo.`create_at`),COUNT(1) FROM panda.`odi_order` oo
    WHERE oo.`order_status` IN (1,2,4,5)
    AND oo.`create_at` > '{}'
    and oo.`create_at` < '{}'
    GROUP BY DATE(oo.`create_at`),HOUR(oo.`create_at`)
    '''
    tomorrow = day + datetime.timedelta(1)
    scur.execute(query_sql.format(day.strftime(date_format), tomorrow.strftime(date_format)))
    result = scur.fetchall()

    hours = np.zeros(24)
    for r in result:
        hours[int(r[1])] = r[2]

    sheet.write(i+1, 0, day.strftime(date_format))
    for j, hour in enumerate(hours):
        sheet.write(i+1, j+1, hour)
    print('runtime: ', time.time() - start_time)
path = cf.get('path', 'path')
wb.save(path + today.strftime(date_format) + 'hoursale.xls')
scur.close()
scon.close()
print('overtime: ', time.time() - start_time)


