import pymysql as db
import datetime
import configparser
import xlwt
import time


class Product:
    def __init__(self, model, props):
        self.model = model
        self.props = props

startTime = time.time()
warehousenums = {1: '分拾', 2: '检测', 3: '市场', 4: '上架', 5: '维修', 6: '报废', 7: 'B端', 8: '预上架', 9: '外包维修',
                 11: '京东', 12: '待卖'}

cf = configparser.ConfigParser()
cf.read('conf.conf')
dbhost = cf.get('db', 'db_host')
dbuser = cf.get('db', 'db_user')
dbport = cf.getint('db', 'db_port')
dbpass = cf.get('db', 'db_pass')
dbase = cf.get('db', 'db_db')
scon = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, charset='utf8')
scur = scon.cursor()

wb = xlwt.Workbook()

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

print('start scanning database...', time.time()-startTime)

for wnum in warehousenums:
    storeSql = '''
    SELECT pm.`model_name`,sw.`key_props`,sw.`warehouse_num` FROM panda.`stg_warehouse` sw
    LEFT JOIN panda.`pdi_model` pm
    ON pm.`model_id` =sw.`model_id`
    WHERE sw.`warehouse_status`=1
    AND sw.`warehouse_num` = {}
    '''
    scur.execute(storeSql.format(wnum))
    stores = scur.fetchall()
    products = []
    for s in stores:
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
            sku[name] = sku[name] + 1
        else:
            sku[name] = 1

    sheet = wb.add_sheet(warehousenums[wnum])
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

    print('runtime...', time.time()-startTime)

today = datetime.datetime.today()
dateFormat = '%Y-%m-%d'

path = cf.get('path', 'path')
wb.save(path + today.strftime(dateFormat) + 'store.xls')

scur.close()
scon.close()
print('overtime...', time.time()-startTime)
