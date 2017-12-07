import pymysql as db
import config as conf
import time
import xlwt

start_time = time.time()

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('sheet')

# cf = conf.test
cf = conf.product
src_con = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], db=cf['db'], charset='utf8')
src_cur = src_con.cursor()

properties_sql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
version_dict = conf.properties_dict(src_cur, properties_sql, 5)

color_dict = conf.properties_dict(src_cur, properties_sql, 10)

memory_dict = conf.properties_dict(src_cur, properties_sql, 11)

rate_dict = conf.properties_dict(src_cur, properties_sql, 12)

# products_sql = '''
# SELECT pp.key_props FROM panda.`pdi_product` pp
# LEFT JOIN panda.`stg_warehouse` sw
# ON sw.imei=pp.tag
# WHERE pp.status =1
# AND sw.`warehouse_status` =1
# '''
products_sql = '''
SELECT pp.key_props FROM panda.`pdi_product` pp
'''

src_cur.execute(products_sql)
products_result = src_cur.fetchall()
product_dict = conf.product_count(products_result, version_dict, memory_dict, color_dict, rate_dict)

sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '成色')
sheet.write(0, 4, '数量')
for i, p in enumerate(product_dict):
    pv, pm, pc, pr = p.split(':')
    sheet.write(i+1, 0, pv)
    sheet.write(i+1, 1, pm)
    sheet.write(i+1, 2, pc)
    sheet.write(i+1, 3, pr)
    sheet.write(i+1, 4, product_dict[p])

workbook.save(conf.path + conf.today.strftime(conf.date_format) + 'spu.xls')

src_cur.close()
src_con.close()
