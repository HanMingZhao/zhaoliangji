import pymysql as db
import config as conf
import time
import xlwt

start_time = time.time()

workbook = xlwt.Workbook()

cf = conf.test
# cf = conf.product
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

products_sql = '''
SELECT pp.key_props FROM panda.pdi_product pp 
LEFT JOIN panda.pdi_model pm 
ON pp.model_id = pm.model_id 
WHERE pm.model_name NOT LIKE '%iphone%'
'''

src_cur.execute(products_sql)
products_result = src_cur.fetchall()
product_dict = conf.product_count(products_result, version_dict, memory_dict, color_dict, rate_dict)

version_color_dict = conf.version_color_dict
version_memory_dict = conf.version_memory_dict
sheet = workbook.add_sheet('iphone')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '成色')
sheet.write(0, 4, '数量')
for rate in rate_dict:
    for version in version_color_dict:
        for color in version_color_dict[version]:
            for memory in version_memory_dict[version]:
                row = len(sheet.rows)
                sheet.write(row, 0, version_dict[str(version)])
                sheet.write(row, 1, memory_dict[str(memory)])
                sheet.write(row, 2, color_dict[str(color)])
                sheet.write(row, 3, rate_dict[rate])
                sheet.write(row, 4, 1)

sheet = workbook.add_sheet('android')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '成色')
sheet.write(0, 4, '数量')
for rate in rate_dict:
    for p in product_dict:
        pv, pm, pc = p.split(':')
        row = len(sheet.rows)
        sheet.write(row, 0, pv)
        sheet.write(row, 1, pm)
        sheet.write(row, 2, pc)
        sheet.write(row, 3, rate_dict[rate])
        sheet.write(row, 4, product_dict[p])

workbook.save(conf.today.strftime(conf.date_format) + 'spu.xls')

src_cur.close()
src_con.close()
