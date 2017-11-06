import pymysql as db
import xlwt
import datetime as dt
import configparser


class Product:
    def __init__(self, model, props, wnum, pvsid):
        self.model = model
        self.props = props
        self.wnum = wnum
        self.pvsid = pvsid

warehouse = {4: '上架', 8: '预上架'}
pset = set()
shangjia = {}
yushangjia = {}

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

for wnum in warehouse:
    storagesql = '''
    SELECT pm.`model_name`,pp.`key_props`,sw.`warehouse_num`, ppv.`pvs_id` FROM panda.`stg_warehouse` sw
    LEFT JOIN panda.`pdi_model` pm
    ON pm.`model_id` =sw.`model_id`
    LEFT JOIN panda.pdi_product pp 
    ON pp.product_id = sw.product_id
    LEFT JOIN panda.`pdi_param_value` ppv
    ON sw.`product_id` = ppv.`product_id`
    WHERE sw.`warehouse_status`=1
    AND ppv.`p_id`=12
    AND sw.`warehouse_num` = {}
    '''
    scur.execute(storagesql.format(wnum))
    storages = scur.fetchall()

    products = []
    for s in storages:
        p = Product(s[0], s[1], s[2], s[3])
        properties = p.props.split(';')
        for f in properties:
            feature = f.split(':')
            if feature[0] == '5':
                p.version = vd[feature[1]]
            if feature[0] == '11':
                p.memory = md[feature[1]]
            if feature[0] == '12':
                p.quality = qd[feature[1]]
        p.battery = bd[str(p.pvsid)]
        products.append(p)

    for prod in products:
        name = prod.version + ':' + prod.memory + ':' + prod.quality + ':' + prod.battery
        pset.add(name)
        if wnum == 4:
            if name in shangjia:
                shangjia[name] = shangjia[name] + 1
            else:
                shangjia[name] = 1
        else:
            if name in yushangjia:
                yushangjia[name] = yushangjia[name] + 1
            else:
                yushangjia[name] = 1

wb = xlwt.Workbook()
sheet1 = wb.add_sheet(warehouse[4])
sheet1.write(0, 0, '型号')
sheet1.write(0, 1, '内存')
sheet1.write(0, 2, '成色')
sheet1.write(0, 3, '电池')
sheet1.write(0, 4, '数量')
for i, s in enumerate(shangjia):
    v, m, q, b = s.split(':')
    sheet1.write(i+1, 0, v)
    sheet1.write(i+1, 1, m)
    sheet1.write(i+1, 2, q)
    sheet1.write(i+1, 3, b)
    sheet1.write(i+1, 4, shangjia[s])

sheet2 = wb.add_sheet(warehouse[8])
sheet2.write(0, 0, '型号')
sheet2.write(0, 1, '内存')
sheet2.write(0, 2, '成色')
sheet2.write(0, 3, '电池')
sheet2.write(0, 4, '数量')
for i, s in enumerate(yushangjia):
    v, m, q, b = s.split(':')
    sheet2.write(i+1, 0, v)
    sheet2.write(i + 1, 1, m)
    sheet2.write(i + 1, 2, q)
    sheet2.write(i + 1, 3, b)
    sheet2.write(i + 1, 4, shangjia[s])

path = conf.get('path', 'path')
today = str(dt.datetime.today().date())
wb.save(path + today + 'storeinfo.xls')

scur.close()
scon.close()
