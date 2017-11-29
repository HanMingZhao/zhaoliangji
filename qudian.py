import pymysql as db
import configparser
import xlwt
import os
import sys
import time
sys.path.append(os.path.abspath(os.path.dirname(__file__)))

stime = time.time()


class Product:
    def __init__(self, name, props):
        self.name = name
        self.props = props

cf = configparser.ConfigParser()
cf.read('/conf.conf')
option = 'db'
dbhost = cf.get(option, 'host')
dbuser = cf.get(option, 'user')
dbport = cf.getint(option, 'port')
dbpass = cf.get(option, 'pass')
dbase = cf.get(option, 'db')
scon = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, port=dbport, charset='utf8')
scur = scon.cursor()

wb = xlwt.Workbook()
sheet = wb.add_sheet('sheet')

modelSql = '''
SELECT pm.`model_name`,pp.`key_props` FROM panda.`odi_order` oo
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.`product_id`
LEFT JOIN panda.`pdi_model` pm
ON pm.`model_id` = pp.`model_id`
WHERE oo.`user_id`=118069
'''

propSql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = {}
scur.execute(propSql.format(5))
versions = scur.fetchall()
for v in versions:
    vd[str(v[0])] = v[1]

cd = {}
scur.execute(propSql.format(10))
colors = scur.fetchall()
for c in colors:
    cd[str(c[0])] = c[1]

md = {}
scur.execute(propSql.format(11))
memories = scur.fetchall()
for m in memories:
    md[str(m[0])] = m[1]

scur.execute(modelSql)
result = scur.fetchall()
products = []
for r in result:
    p = Product(r[0], r[1])
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
        sku[name] = sku[name] + 1
    else:
        sku[name] = 1

sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '总库存')

for i, product in enumerate(sku):
    v, m, c = product.split(':')
    sheet.write(i + 1, 0, v)
    sheet.write(i + 1, 1, m)
    sheet.write(i + 1, 2, c)
    sheet.write(i + 1, 3, sku[product])

path = cf.get('path', 'path')
wb.save(path + 'qudian.xls')

scur.close()
scon.close()
print('overtime...', time.time()-stime)
