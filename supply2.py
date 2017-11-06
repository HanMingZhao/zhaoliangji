import pymysql as db
import xlwt
from datetime import datetime
import configparser


class Product:
    def __init__(self, model, props, wnum):
        self.model = model
        self.props = props
        self.wnum = wnum

cf = configparser.ConfigParser()
dbhost = cf.get('db', 'db_host')
dbuser = cf.get('db', 'db_user')
dbport = cf.getint('db', 'db_port')
dbpass = cf.get('db', 'db_pass')
dbase = cf.get('db', 'db_db')
scon = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, charset='utf8')
scur = scon.cursor()

warehouse_num = [1, 2, 4, 8]
pset = set()
fenshi = {}
jiance = {}
shangjia = {}

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

for wnum in warehouse_num:
    storagesql = '''
    SELECT pm.`model_name`,sw.`key_props`,sw.`warehouse_num` FROM panda.`stg_warehouse` sw
    LEFT JOIN panda.`pdi_model` pm
    ON pm.`model_id` =sw.`model_id`
    WHERE sw.`warehouse_status`=1
    AND sw.`warehouse_num` IN ({})
    '''
    scur.execute(storagesql.format(wnum))
    storages = scur.fetchall()

    products = []
    for s in storages:
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
        products.append(p)

    for prod in products:
        name = prod.version + ':' + prod.memory + ':' + prod.color
        pset.add(name)
        if wnum == 1:
            if name in fenshi:
                fenshi[name] = fenshi[name] + 1
            else:
                fenshi[name] = 1
        elif wnum == 2:
            if name in jiance:
                jiance[name] = jiance[name] + 1
            else:
                jiance[name] = 1
        else:
            if name in shangjia:
                shangjia[name] = shangjia[name] + 1
            else:
                shangjia[name] = 1

wb = xlwt.Workbook()
sheet = wb.add_sheet('sheet1')
sheet.write(0, 0, '日期')
sheet.write(0, 1, '型号')
sheet.write(0, 2, '内存')
sheet.write(0, 3, '颜色')
sheet.write(0, 4, '分拾库')
sheet.write(0, 5, '检测库')
sheet.write(0, 6, '上架/预上架')
sheet.write(0, 7, '总库存')

today = str(datetime.today().date())
for i, product in enumerate(pset):
    v, m, c = product.split(':')
    sheet.write(i+1, 0, today)
    sheet.write(i+1, 1, v)
    sheet.write(i+1, 2, m)
    sheet.write(i+1, 3, c)
    sheet.write(i+1, 4, fenshi[product] if product in fenshi else 0)
    sheet.write(i+1, 5, jiance[product] if product in jiance else 0)
    sheet.write(i+1, 6, shangjia[product] if product in shangjia else 0)
    sheet.write(i+1, 7, xlwt.Formula('SUM(e{}:g{})'.format(i+2, i+2)))
path = cf.get('path', 'path')
wb.save(path + today + 'supply.xls')

scur.close()
scon.close()
