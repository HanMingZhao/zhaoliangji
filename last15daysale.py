import pymysql as db
import config as conf
import datetime
import xlwt

cf = conf.product
connect = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cursor = connect.cursor()
wb = xlwt.Workbook()

propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = conf.properties_dict(cursor, propsql, 5)
md = conf.properties_dict(cursor, propsql, 11)
cd = conf.properties_dict(cursor, propsql, 10)

prop_sql = '''
SELECT pp.`key_props` FROM panda.`odi_order` oo
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.`product_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`pay_at`>'{}'
AND oo.`pay_at`<'{}'
'''
fifteen_day = conf.today-datetime.timedelta(15)
cursor.execute(prop_sql.format(fifteen_day.strftime(conf.date_format), conf.today.strftime(conf.date_format)))
grounding_result = cursor.fetchall()
grounding_dict = conf.product_count(grounding_result, vd, md, cd)
sheet = wb.add_sheet('sheet')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '数量')
for i, r in enumerate(grounding_dict):
    pv, pm, pc = r.split(':')
    sheet.write(i+1, 0, pv)
    sheet.write(i+1, 1, pm)
    sheet.write(i+1, 2, pc)
    sheet.write(i+1, 3, grounding_dict[r])

wb.save(conf.path + '15daysale.xls')
cursor.close()
connect.close()
