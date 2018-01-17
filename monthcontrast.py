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
for r in result2017:
    row = len(sheet.rows)
    pv, pm, pc = r.split(':')
    sheet.write(row, 0, pv)
    sheet.write(row, 1, pm)
    sheet.write(row, 2, pc)
    sheet.write(row, 3, result2017[r])

add18 = []

sheet = wb.add_sheet('1yue')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '数量')
for r in result2018:
    row = len(sheet.rows)
    pv, pm, pc = r.split(':')
    sheet.write(row, 0, pv)
    sheet.write(row, 1, pm)
    sheet.write(row, 2, pc)
    sheet.write(row, 3, result2018[r])
    if r not in result2017:
        add18.append(r)

sheet = wb.add_sheet('xinzeng')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
for a in add18:
    pv, pm, pc = a.split(':')
    sheet.write(row, 0, pv)
    sheet.write(row, 1, pm)
    sheet.write(row, 2, pc)

wb.save('shangjiaxinzeng.xls')
conf.product_cursor.close()
conf.product_connect.close()
