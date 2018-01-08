import pymysql as db
import config as conf
import xlwt
import datetime

cf = conf.product
connect = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cursor = connect.cursor()
wb = xlwt.Workbook()
sheet = wb.add_sheet('sheet')
start_time = datetime.datetime.strptime('2017-9-28', conf.date_format)
end_time = start_time + datetime.timedelta(14)
sheet.write(0, 0, '周期')
sheet.write(0, 1, '下单')
sheet.write(0, 2, '下周购买')
sheet.write(0, 3, '转化')
while end_time < conf.today:
    next_time = start_time + datetime.timedelta(7)
    count_sql = '''
    SELECT oo.`user_id` FROM panda.`odi_order` oo
    WHERE oo.`order_status` IN (0,3)
    AND oo.`create_at` >'{}'
    AND oo.`create_at` < '{}'
    GROUP BY oo.`user_id`
    HAVING COUNT(1)>1
    '''
    count = cursor.execute(count_sql.format(start_time.strftime(conf.date_format),
                                            next_time.strftime(conf.date_format)))
    buy_sql = '''
    SELECT COUNT(1) FROM panda.`odi_order` oo
    WHERE oo.`order_status` IN (1,2,4,5)
    AND oo.`order_type` IN (1,2)
    AND oo.`user_id` IN (
    SELECT oo.`user_id` FROM panda.`odi_order` oo
    WHERE oo.`order_status` IN (0,3)
    AND oo.`create_at` >'{}'
    AND oo.`create_at` < '{}'
    GROUP BY oo.`user_id`
    HAVING COUNT(1)>1
    )
    AND oo.`pay_at`>'{}'
    AND oo.`pay_at`<'{}'
    '''
    end_time = next_time + datetime.timedelta(7)
    cursor.execute(buy_sql.format(start_time.strftime(conf.date_format), next_time.strftime(conf.date_format),
                                  next_time.strftime(conf.date_format), end_time.strftime(conf.date_format)))
    result = cursor.fetchone()
    row = len(sheet.rows)
    sheet.write(row, 0, start_time.strftime(conf.date_format))
    sheet.write(row, 1, count)
    sheet.write(row, 2, result[0])
    sheet.write(row, 3, result[0]/count)

wb.save(conf.path+'nwb.xls')
cursor.close()
connect.close()
