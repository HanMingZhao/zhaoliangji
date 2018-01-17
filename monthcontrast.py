import config as conf
import xlwt

propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = conf.properties_dict(conf.product_cursor, propsql, 5)
md = conf.properties_dict(conf.product_cursor, propsql, 11)
cd = conf.properties_dict(conf.product_cursor, propsql, 10)

shangjia_sql = '''
SELECT ppt.`key_props` FROM panda.`pdi_product_track` ppt
WHERE ppt.`track_type`=1
AND ppt.`product_status`!=3
AND ppt.`created_at`>'{}' 
{}
'''
conf.product_cursor.execute(shangjia_sql.format('2018-1-1', ''))
result2018 = conf.product_cursor.fetchall()
shangjia2018dict = conf.product_count(result2018, vd, md, cd)

conf.product_cursor.execute(shangjia_sql.format('2017-12-1', 'and ppt.created_at < \'2018-1-1\''))
result2017 = conf.product_cursor.fetchall()
shangjia2017dict = conf.product_count(result2017, vd, md, cd)

wb = xlwt.Workbook()
sheet = wb.add_sheet("12yue")
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '数量')
for r in shangjia2017dict:
    row = len(sheet.rows)
    pv, pm, pc = r.split(':')
    sheet.write(row, 0, pv)
    sheet.write(row, 1, pm)
    sheet.write(row, 2, pc)
    sheet.write(row, 3, shangjia2017dict[r])

add18 = []

sheet = wb.add_sheet('1yue')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '数量')
for r in shangjia2018dict:
    row = len(sheet.rows)
    pv, pm, pc = r.split(':')
    sheet.write(row, 0, pv)
    sheet.write(row, 1, pm)
    sheet.write(row, 2, pc)
    sheet.write(row, 3, shangjia2018dict[r])
    if r not in shangjia2017dict:
        add18.append(r)
sale2018 = '''
SELECT pp.key_props FROM panda.`odi_order` oo
LEFT JOIN panda.`pdi_product` pp
ON pp.product_id = oo.product_id
WHERE oo.order_status IN (1,2,4,5)
AND oo.order_type IN (1,2)
AND oo.pay_at >'2018-1-1'
'''
count = conf.product_cursor.execute(sale2018)
sale2018result = conf.product_cursor.fetchall()
sale2018dict = conf.product_count(sale2018result, vd, md, cd)
sheet = wb.add_sheet('xinzeng')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '销量')
sheet.write(0, 4, '占比')
sheet.write(0, 8, '总销量')
sheet.write(1, 8, count)
for a in add18:
    row = len(sheet.rows)
    pv, pm, pc = a.split(':')
    sheet.write(row, 0, pv)
    sheet.write(row, 1, pm)
    sheet.write(row, 2, pc)
    sheet.write(row, 3, sale2018dict[a] if a in sale2018dict else 0)
    sheet.write(row, 4, sale2018dict[a]/count if a in sale2018dict else 0)

wb.save('shangjiaxinzeng.xls')
conf.product_cursor.close()
conf.product_connect.close()
