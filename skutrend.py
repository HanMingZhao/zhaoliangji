import config as conf
import xlwt
import datetime

wb = xlwt.Workbook()
sheet = wb.add_sheet('sheet')
sheet.write(0, 0, '机型')
# sheet.write(0, 1, '日期')
# sheet.write(0, 2, '销量')
each_day_sql = '''
SELECT DATE(oo.pay_at),COUNT(1) FROM panda.`odi_order` oo
LEFT JOIN panda.`pdi_product` pp
ON oo.product_id = pp.`product_id`
WHERE oo.order_status IN (1,2,4,5)
AND oo.order_type IN (1,2)
AND pp.`key_props` LIKE '%{};%'
AND pp.`key_props` LIKE '%{};%'
AND pp.`key_props` LIKE '%{};%'
AND oo.pay_at > '{}'
AND oo.pay_at < '{}'
GROUP BY DATE(oo.pay_at)
'''
last_15_day = conf.today-datetime.timedelta(15)

iphone_top_sql = '''
SELECT sws.key_props,sws.sku_name FROM panda.`stg_warning_sku` sws
WHERE sws.category = 1
AND sws.brand_id = 1
AND sws.type_id =1
'''
conf.product_cursor.execute(iphone_top_sql)
iphone_top_result = conf.product_cursor.fetchall()
for i in range(15):
    day = last_15_day+datetime.timedelta(i)
    sheet.write(0, i+1, day.strftime(conf.date_format))
for itr in iphone_top_result:
    p = itr[0].split(';')
    conf.product_cursor.execute(each_day_sql.format(p[0], p[1], p[2], last_15_day.strftime(conf.date_format),
                                                    conf.today.strftime(conf.date_format)))
    each_day_result = conf.product_cursor.fetchall()
    sale_dict = {}
    for x in range(15):
        day = last_15_day + datetime.timedelta(x)
        sale_dict[day.strftime(conf.date_format)] = 0
    for edr in each_day_result:
        sale_dict[edr[0].strftime(conf.date_format)] = edr[1]
    row = len(sheet.rows)
    sheet.write(row, 0, itr[1])
    for i, sd in enumerate(sale_dict):
        sheet.write(row, i+1, sale_dict[sd])
wb.save(conf.path + 'trendsku.xls')
conf.product_cursor.close()
conf.product_connect.close()
