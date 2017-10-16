import pymysql as db
import xlwt
from datetime import datetime


class Product:
    def __init__(self, model, props):
        self.model = model
        self.props = props

scon = db.connect(host='rm-bp13wnvyc2dh86ju1.mysql.rds.aliyuncs.com', user='panda_reader', passwd='zhaoliangji3503',
                  db='panda', charset='utf8')
scur = scon.cursor()

storagesql = '''
SELECT pm.`model_name`,sw.`key_props`,sw.`warehouse_num` FROM panda.`stg_warehouse` sw
LEFT JOIN panda.`pdi_model` pm
ON pm.`model_id` =sw.`model_id`
WHERE sw.`warehouse_status`=1
AND sw.`warehouse_num` IN (1,2,4,8)
'''
scur.execute(storagesql)
storages = scur.fetchall()

versionsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = 5
'''
vd = {}
scur.execute(versionsql)
versions = scur.fetchall()
for v in versions:
    vd[str(v[0])] = v[1]

colorsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = 10
'''
cd = {}
scur.execute(colorsql)
colors = scur.fetchall()
for c in colors:
    cd[str(c[0])] = c[1]

memorysql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = 11
'''
md = {}
scur.execute(memorysql)
memorys = scur.fetchall()
for m in memorys:
    md[str(m[0])] = m[1]

products = []
for s in storages:
    p = Product(s[0], s[1])
    properties = p.props.split(';')
    for f in properties:
        feature = f.split(':')
        if feature[0] == '5':
            p.version = vd[feature[1]]
        if feature[0] == '10':
            p.color = cd[feature[1]]
        if feature[0] == '11':
            p.memory = md[feature[1]]
    products.append(p)

sku = {}
for prod in products:
    name = prod.version + ':' + prod.memory + ':' + prod.color
    if name in sku:
        sku[name] = sku[name]+1
    else:
        sku[name] = 1

wb = xlwt.Workbook()
sheet = wb.add_sheet('sheet1')
sheet.write(0, 0, '日期')
sheet.write(0, 1, '型号')
sheet.write(0, 2, '内存')
sheet.write(0, 3, '颜色')
sheet.write(0, 4, '总库存')

today = str(datetime.today().date())
for i, product in enumerate(sku):
    v, m, c = product.split(':')
    sheet.write(i+1, 0, today)
    sheet.write(i+1, 1, v)
    sheet.write(i+1, 2, m)
    sheet.write(i+1, 3, c)
    sheet.write(i+1, 4, sku[product])

wb.save(today + 'supply.xls')

scur.close()
scon.close()