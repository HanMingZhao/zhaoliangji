import config as conf
import datetime
import xlwt

base_sql = '''
SELECT DISTINCT(oo.`user_id`) FROM panda.`odi_order` oo
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.order_type IN (1,2)
AND oo.`pay_at` >'{}'
AND oo.`pay_at` <'{}'
{}
'''

condition = '''
AND oo.`user_id` IN 
(
SELECT DISTINCT(oo.`user_id`) FROM panda.`odi_order` oo
WHERE oo.`order_status` NOT IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`create_at` > '{}'
AND oo.`create_at` < '{}'
)
'''
wb = xlwt.Workbook()
sheet = wb.add_sheet('sheet')
sheet.write(0, 0, '日期')
sheet.write(0, 1, '总量')
sheet.write(0, 2, '延迟量')
sheet.write(0, 3, '比例')
start = datetime.datetime.strptime('2017-11-1', conf.date_format)
for i in range(100):
    day = start + datetime.timedelta(i)
    if day >= conf.today:
        break
    next_day = day + datetime.timedelta(1)
    before_day = day - datetime.timedelta(2)
    total = conf.product_cursor.execute(base_sql.format(day.strftime(conf.date_format),
                                                        next_day.strftime(conf.date_format), ''))
    past = conf.product_cursor.execute(base_sql.format(day.strftime(conf.date_format),
                                                       next_day.strftime(conf.date_format),
                                                       condition.format(before_day.strftime(conf.date_format),
                                                                        day.strftime(conf.date_format))))
    row = len(sheet.rows)
    sheet.write(row, 0, day.strftime(conf.date_format))
    sheet.write(row, 1, total)
    sheet.write(row, 2, past)
    sheet.write(row, 3, past/total)

wb.save('delay.xls')
conf.product_cursor.close()
conf.product_connect.close()
