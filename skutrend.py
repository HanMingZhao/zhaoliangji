import config as conf
import xlwt
import datetime

wb = xlwt.Workbook()

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

sku_trend_sql = '''
SELECT sws.key_props,sws.sku_name FROM panda.`stg_warning_sku` sws
WHERE sws.category = {}
AND sws.brand_id  {}
AND sws.type_id ={}
'''


def sku_trend_query(sheet_name, category, brand_id, type_id):
    conf.product_cursor.execute(sku_trend_sql.format(category, brand_id, type_id))
    iphone_top_result = conf.product_cursor.fetchall()
    sheet = wb.add_sheet(sheet_name)
    sheet.write(0, 0, '机型')
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

sku_trend_query('iphone爆款', 2, '=1', 1)
sku_trend_query('iphone主要', 1, '=1', 1)
sku_trend_query('ipad爆款', 2, '=1', 2)
sku_trend_query('ipad主要', 1, '=1', 2)
sku_trend_query('android爆款', 2, '!=1', 1)
sku_trend_query('android主要', 1, '!=1', 1)
wb.save(conf.path + 'trendsku.xls')
conf.product_cursor.close()
conf.product_connect.close()
