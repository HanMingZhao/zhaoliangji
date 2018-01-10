import pymysql as db
import config as conf
import datetime
import xlwt

cf = conf.product
connect = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cursor = connect.cursor()
wb = xlwt.Workbook()

sku_sql = '''
SELECT sws.`key_props`,sws.`sku_name` FROM panda.`stg_warning_sku` sws
LEFT JOIN panda.`pdi_model` pm
ON sws.`model_id` = pm.`model_id`
WHERE sws.`category` IN (1,2)
AND pm.`model_name` LIKE '%iphone%'
'''
cursor.execute(sku_sql)
result = cursor.fetchall()
sku_dict = {}
for r in result:
    sku_dict[r[0]] = r[1]

sheet = wb.add_sheet('sheet')
sheet.write(0, 0, '日期')
sheet.write(0, 1, 'sku')
sheet.write(0, 2, '数量')
last_15_day = conf.today - datetime.timedelta(15)
for sd in sku_dict:
    daily_sql = '''
    SELECT DATE(oo.pay_at),COUNT(1) FROM panda.`odi_order` oo
    LEFT JOIN panda.`pdi_product` pp
    ON oo.product_id = pp.product_id
    WHERE oo.order_status IN (1,2,4,5)
    AND oo.order_type IN (1,2)
    AND oo.pay_at >'{}'
    AND pp.key_props LIKE '%{};%'
    AND pp.key_props LIKE '%{};%'
    AND pp.key_props LIKE '%{};%'
    GROUP BY DATE(oo.pay_at)
    '''
    t = sd.split(';')
    cursor.execute(daily_sql.format(last_15_day.strftime(conf.date_format), t[0], t[1], t[2]))
    daily_result = cursor.fetchall()
    for dr in daily_result:
        row = len(sheet.rows)
        sheet.write(0, 0, dr[0].strftime(conf.date_format))
        sheet.write(0, 1, sku_dict[sd])
        sheet.write(0, 2, dr[1])

wb.save(conf.path + '15daysale.xls')
cursor.close()
connect.close()
