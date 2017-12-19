import pymysql
import config
import xlwt
import time

start_time = time.time()


class Product:
    def __init__(self, props):
        self.props = props
cf = config.product
src_con = pymysql.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], db=cf['db'],
                          charset=config.charset)
src_cur = src_con.cursor()

workbook = xlwt.Workbook()

props_sql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
version_dict = config.properties_dict(src_cur, props_sql, 5)
memory_dict = config.properties_dict(src_cur, props_sql, 11)
color_dict = config.properties_dict(src_cur, props_sql, 10)

print('start...', time.time()-start_time)
sale_sql = '''
SELECT pp.`key_props` FROM panda.`odi_order` oo
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.`product_id`
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id`= pm.`model_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`pay_at` > '2017-5-1'
'''
src_cur.execute(sale_sql)
result = src_cur.fetchall()
product_dict = config.product_count(result, version_dict, memory_dict, color_dict)
sheet = workbook.add_sheet('sheet')
config.sheet_head(sheet, 0)
config.write_sheet1(product_dict, sheet, 0)

workbook.save('预备.xls')
src_cur.close()
src_con.close()
print('over...', time.time()-start_time)
