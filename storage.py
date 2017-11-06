import pymysql as db
import xlwt
from datetime import datetime
import configparser


class Product:
    def __init__(self, model, props, times):
        self.model = model
        self.props = props
        self.times = times

cf = configparser.ConfigParser()
dbhost = cf.get('db', 'db_host')
dbuser = cf.get('db', 'db_user')
dbport = cf.getint('db', 'db_port')
dbpass = cf.get('db', 'db_pass')
dbase = cf.get('db', 'db_db')
scon = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, charset='utf8')
scur = scon.cursor()

today = str(datetime.today().date())
path = cf.get('path', 'path')
warehouse_num = [1, 2, 4, 8]
warehouse_dict = {1: '分拾', 2: '检测', 4: '上架', 8: '预上架'}
pset = set()
fenshigt7 = {}
fenshigt15 = {}
fenshigt30 = {}
jiancegt7 = {}
jiancegt15 = {}
jiancegt30 = {}
shangjiagt7 = {}
shangjiagt15 = {}
shangjiagt30 = {}
yushangjiagt7 = {}
yushangjiagt15 = {}
yushangjiagt30 = {}

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

wbook = xlwt.Workbook()

for wnum in warehouse_num:
    storagesql = '''
     SELECT pm.model_name,psw.key_props,
    IF(psw.change_time>0,(UNIX_TIMESTAMP(NOW())-UNIX_TIMESTAMP(psw.change_time))/60/60/24,
    (UNIX_TIMESTAMP(NOW())-UNIX_TIMESTAMP(psw.in_time))/60/60/24) `times` FROM
    (
    SELECT sw.*,pp.`buy_at` FROM panda.`stg_warehouse` sw
    LEFT JOIN panda.`pdi_product` pp ON sw.`product_id` = pp.`product_id`
    ) psw
    LEFT JOIN panda.`pdi_model`  pm  ON  psw.model_id = pm.model_id
    WHERE psw.warehouse_status = 1
    AND psw.warehouse_num = {}
    ORDER BY pm.model_name ,times,imei
    '''
    scur.execute(storagesql.format(wnum))
    storages = scur.fetchall()

    products = []
    for s in storages:
        product = Product(s[0], s[1], s[2])
        properties = product.props.split(';')
        for f in properties:
            feature = f.split(':')
            if feature[0] == '5':
                product.version = vd[feature[1]]
            if feature[0] == '10':
                product.color = cd[feature[1]]
            if feature[0] == '11':
                product.memory = md[feature[1]]
        products.append(product)

    for prod in products:
        name = prod.version + ':' + prod.memory + ':' + prod.color
        pset.add(name)
        if wnum == 1:
            if prod.times > 30:
                if name in fenshigt30:
                    fenshigt30[name] = fenshigt30[name] + 1
                else:
                    fenshigt30[name] = 1
            if 15 < prod.times < 30:
                if name in fenshigt15:
                    fenshigt15[name] = fenshigt15[name] + 1
                else:
                    fenshigt15[name] = 1
            if 7 < prod.times < 15:
                if name in fenshigt7:
                    fenshigt7[name] = fenshigt7[name] + 1
                else:
                    fenshigt7[name] = 1
        if wnum == 2:
            if 7 < prod.times < 15:
                if name in jiancegt7:
                    jiancegt7[name] = jiancegt7[name] + 1
                else:
                    jiancegt7[name] = 1
            if 15 < prod.times < 30:
                if name in jiancegt15:
                    jiancegt15[name] = jiancegt15[name] + 1
                else:
                    jiancegt15[name] = 1
            if prod.times > 30:
                if name in jiancegt30:
                    jiancegt30[name] = jiancegt30[name] + 1
                else:
                    jiancegt30[name] = 1
        if wnum == 4:
            if 7 < prod.times < 15:
                if name in shangjiagt7:
                    shangjiagt7[name] = shangjiagt7[name] + 1
                else:
                    shangjiagt7[name] = 1
            if 15 < prod.times < 30:
                if name in shangjiagt15:
                    shangjiagt15[name] = shangjiagt15[name] + 1
                else:
                    shangjiagt15[name] = 1
            if prod.times > 30:
                if name in shangjiagt30:
                    shangjiagt30[name] = shangjiagt30[name] + 1
                else:
                    shangjiagt30[name] = 1
        if wnum == 8:
            if 7 < prod.times < 15:
                if name in yushangjiagt7:
                    yushangjiagt7[name] = yushangjiagt7[name] + 1
                else:
                    yushangjiagt7[name] = 1
            if 15 < prod.times < 30:
                if name in yushangjiagt15:
                    yushangjiagt15[name] = yushangjiagt15[name] + 1
                else:
                    yushangjiagt15[name] = 1
            if prod.times > 30:
                if name in yushangjiagt30:
                    yushangjiagt30[name] = yushangjiagt30[name] + 1
                else:
                    yushangjiagt30[name] = 1

    sheet = wbook.add_sheet(warehouse_dict[wnum])
    sheet.write(0, 0, '型号')
    sheet.write(0, 1, '内存')
    sheet.write(0, 2, '颜色')
    sheet.write(0, 3, '大于7天')
    sheet.write(0, 4, '大于15天')
    sheet.write(0, 5, '大于30天')

    for i, product in enumerate(pset):
        v, m, c = product.split(':')
        sheet.write(i+1, 0, v)
        sheet.write(i+1, 1, m)
        sheet.write(i+1, 2, c)
        if wnum == 1:
            sheet.write(i+1, 3, fenshigt7[product] if product in fenshigt7 else 0)
            sheet.write(i+1, 4, fenshigt15[product] if product in fenshigt15 else 0)
            sheet.write(i+1, 5, fenshigt30[product] if product in fenshigt30 else 0)
        if wnum == 2:
            sheet.write(i+1, 3, jiancegt7[product] if product in jiancegt7 else 0)
            sheet.write(i+1, 4, jiancegt15[product] if product in jiancegt15 else 0)
            sheet.write(i+1, 5, jiancegt30[product] if product in jiancegt30 else 0)
        if wnum == 4:
            sheet.write(i+1, 3, shangjiagt7[product] if product in shangjiagt7 else 0)
            sheet.write(i+1, 4, shangjiagt15[product] if product in shangjiagt15 else 0)
            sheet.write(i+1, 5, shangjiagt30[product] if product in shangjiagt30 else 0)
        if wnum == 8:
            sheet.write(i + 1, 3, yushangjiagt7[product] if product in yushangjiagt7 else 0)
            sheet.write(i + 1, 4, yushangjiagt15[product] if product in yushangjiagt15 else 0)
            sheet.write(i + 1, 5, yushangjiagt30[product] if product in yushangjiagt30 else 0)

wbook.save(path + today + 'storage.xls')

scur.close()
scon.close()
