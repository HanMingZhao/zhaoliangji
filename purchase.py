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
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`pay_at` > '2017-11-1'
'''
src_cur.execute(sale_sql)
result = src_cur.fetchall()
sale_dict = config.product_count(result, version_dict, memory_dict, color_dict)
sheet = workbook.add_sheet('sheet')

# store_sql = '''
# SELECT sw.`key_props` FROM panda.`stg_warehouse` sw
# LEFT JOIN panda.`pdi_model` pm
# ON pm.`model_id` = sw.`model_id`
# WHERE sw.`warehouse_status`=1
# and sw.warehouse_num in (1,2,4,7)
# '''
store_sql = '''
SELECT pp.key_props FROM panda.`stg_warehouse` sw
LEFT JOIN panda.`pdi_product` pp
ON sw.`product_id` = pp.product_id
WHERE sw.`warehouse_status` = 1
and pp.status = 1
AND sw.`warehouse_num` IN (1,2,4,7)
'''

src_cur.execute(store_sql)
result = src_cur.fetchall()
store_dict = config.product_count(result, version_dict, memory_dict, color_dict)
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '售卖')
sheet.write(0, 4, '库存')
# product_set = set()
# for p in sale_dict:
#     product_set.add(p)
# for p in store_dict:
#     product_set.add(p)
for i, p in enumerate(sale_dict):
    pv, pm, pc = p.split(':')
    sheet.write(i+1, 0, pv)
    sheet.write(i+1, 1, pm)
    sheet.write(i+1, 2, pc)
    sheet.write(i+1, 3, sale_dict[p] if p in sale_dict else 0)
    sheet.write(i+1, 4, store_dict[p] if p in store_dict else 0)

workbook.save(config.path + config.today.strftime(config.date_format) + 'yubei.xls')
src_cur.close()
src_con.close()
print('over...', time.time()-start_time)
