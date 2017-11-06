import pymysql as db
import xlwt
import datetime as dt
import configparser


class Product:
    def __init__(self, model, props, pvsid):
        self.model = model
        self.props = props
        self.pvsid = pvsid

conf = configparser.ConfigParser()
conf.read('conf.conf')
dbhost = conf.get('db', 'db_host')
dbuser = conf.get('db', 'db_user')
dbport = conf.getint('db', 'db_port')
dbpass = conf.get('db', 'db_pass')
dbase = conf.get('db', 'db_db')
scon = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, charset='utf8')
scur = scon.cursor()

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

qualitysql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = 12
'''
qd = {}
scur.execute(qualitysql)
qualities = scur.fetchall()
for q in qualities:
    qd[str(q[0])] = q[1]

batterysql = '''
SELECT ppv.id,ppv.`p_values` FROM panda.`pdi_param_values` ppv
WHERE ppv.`p_id`=12
'''
bd = {}
scur.execute(batterysql)
batteries = scur.fetchall()
for b in batteries:
    bd[str(b[0])] = b[1]

salesql = '''
SELECT pm.`model_name`,pp.`key_props`,ppv.`pvs_id` FROM panda.`odi_order` oo
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.`product_id`
LEFT JOIN panda.`pdi_param_value` ppv
ON oo.`product_id` = ppv.`product_id`
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id` = pm.`model_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`create_at` > '2017-09-25'
AND ppv.`p_id`=12
'''
scur.execute(salesql)
saleinfo = scur.fetchall()

sales = {}

products = []
for s in saleinfo:
    p = Product(s[0], s[1], s[2])
    properties = p.props.split(';')
    for f in properties:
        feature = f.split(':')
        if feature[0] == '5':
            p.version = vd[feature[1]]
        if feature[0] == '10':
            p.color = cd[feature[1]]
        if feature[0] == '11':
            p.memory = md[feature[1]]
        if feature[0] == '12':
            p.quality = qd[feature[1]]
    p.battery = bd[p.pvsid]
    products.append(p)

for prod in products:
    name = prod.version + ':' + prod.memory + ':' + prod.quality + ':' + prod.battery
    if name in sales:
        sales[name] = sales[name] + 1
    else:
        sales[name] = 1

wb = xlwt.Workbook()
sheet = wb.add_sheet('sheet1')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '成色')
sheet.write(0, 3, '电池')
sheet.write(0, 4, '数量')
today = str(dt.datetime.today().date())
for i, s in enumerate(sales):
    v, m, q, b = s.split(':')
    sheet.write(i+1, 0, v)
    sheet.write(i+1, 1, m)
    sheet.write(i+1, 2, q)
    sheet.write(i+1, 3, b)
    sheet.write(i+1, 4, sales[s])

path = conf.get('path', 'path')
wb.save(path + today + 'salesinfo.xls')

scur.close()
scon.close()
