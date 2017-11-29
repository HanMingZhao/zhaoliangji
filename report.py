import pymysql as db
import xlwt
import datetime
import configparser
import time
import os


class Product:
    def __init__(self, model, props):
        self.model = model
        self.props = props


def product_count(sql_results):
    productlst = []
    for sr in sql_results:
        if sr[0] != None:
            p = Product(sr[0], sr[1])
            properties = p.props.split(';')
            for f in properties:
                feature = f.split(':')
                if feature[0] == '5':
                    p.version = vd[feature[1]]
                if feature[0] == '10':
                    p.color = cd[feature[1]]
                if feature[0] == '11':
                    p.memory = md[feature[1]]
            productlst.append(p)
    productdict = {}
    for prod in productlst:
        name = prod.version + ':' + prod.memory + ':' + prod.color
        if name in productdict:
            productdict[name] = productdict[name] + 1
        else:
            productdict[name] = 1
    return productdict


def write_sheet1(count_dict, sheety, idx):
    for n, obj in enumerate(count_dict):
        objv, objm, objc = obj.split(':')
        sheety.write(n+1, idx, objv)
        sheety.write(n+1, idx+1, objm)
        sheety.write(n+1, idx+2, objc)
        sheety.write(n+1, idx+3, count_dict[obj])


def write_sheet2(sql_result, sheety, idx):
    number = 0
    for n, sr in enumerate(sql_result):
        srdate = sr[0]
        sheety.write(n+1, idx, str(srdate))
        sheety.write(n+1, idx+1, sr[1])
        number += sr[1]
    row = len(sql_result)+1
    sheety.write(row, idx, '总计')
    sheety.write(row, idx+1, number)
    return number


stime = time.time()

cf = configparser.ConfigParser()
cf.read(os.path.dirname(__file__) + '/conf.conf')
option = 'test'
dbhost = cf.get(option, 'host')
dbuser = cf.get(option, 'user')
dbport = cf.getint(option, 'port')
dbpass = cf.get(option, 'pass')
dbase = cf.get(option, 'db')
# dbhost = test['host']
# dbuser = test['user']
# dbpass = test['pass']
# dbase = test['db']
# dbport = test['port']
scon = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, port=dbport, charset='utf8')
scur = scon.cursor()

propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = {}
scur.execute(propsql.format(5))
versions = scur.fetchall()
for v in versions:
    vd[str(v[0])] = v[1]

cd = {}
scur.execute(propsql.format(10))
colors = scur.fetchall()
for c in colors:
    cd[str(c[0])] = c[1]

md = {}
scur.execute(propsql.format(11))
memories = scur.fetchall()
for m in memories:
    md[str(m[0])] = m[1]

print('start scanning database...', time.time()-stime)

dateFormat = '%Y-%m-%d'

wb = xlwt.Workbook()

today = datetime.datetime.strptime('2016-11-28', dateFormat)
#today = datetime.datetime.today()
yesterday = today-datetime.timedelta(1)
month = today.month
year = today.year

first = datetime.datetime.strptime(str(year)+'-'+str(month)+'-'+str(1), dateFormat)

print('日销量...', time.time()-stime)
sheet = wb.add_sheet('销售总计')
sheet.write(0, 0, '日期')
sheet.write(0, 1, '销量')
daySaleSql = '''
SELECT DATE(oo.`pay_at`),COUNT(1) FROM panda.`odi_order` oo
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`pay_at` > '{}'
AND oo.`pay_at` < '{}'
GROUP BY DATE(oo.`pay_at`)
'''

scur.execute(daySaleSql.format(first.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchall()
saleSum = write_sheet2(result, sheet, 0)
target = cf.getint(option, 'target')
sheet.write(len(sheet.rows), 0, '距离目标还差 {} 台'.format(target-saleSum))

print('取消订单数...', time.time()-stime)
sheet.write(0, 5, '日期')
sheet.write(0, 6, '取消订单数')
cancelOrderSql = '''
SELECT t.date,COUNT(1) FROM 
(
SELECT DATE(oo.`create_at`) `date`,COUNT(1) `count` FROM panda.`odi_order` oo
WHERE oo.`order_status` = 3
AND oo.`order_type` IN (1,2)
AND oo.`create_at` > '{}'
AND oo.`create_at` < '{}'
AND oo.`user_id` NOT IN (
SELECT DISTINCT(oo.user_id) FROM panda.`odi_order` oo
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`pay_at` > '{}'
)
GROUP BY DATE(oo.`create_at`),oo.`user_id`
)t
GROUP BY t.date
'''
scur.execute(cancelOrderSql.format(first.strftime(dateFormat), today.strftime(dateFormat), first.strftime(dateFormat)))
result = scur.fetchall()
write_sheet2(result, sheet, 5)

print('退货数...', time.time()-stime)
aftersaleSql = '''
SELECT a.date,a.count `apply`,b.count `finish` FROM (
SELECT DATE(ooa.`created_at`) `date`,COUNT(1) `count` FROM panda.`odi_order_aftersale` ooa
WHERE ooa.`type`=1
AND ooa.`created_at` >'{}'
AND ooa.`created_at` <'{}'
GROUP BY DATE(ooa.`created_at`)
) a LEFT JOIN 
(
SELECT DATE(ooa.`finsh_time`) `date`,COUNT(1) `count` FROM panda.`odi_order_aftersale` ooa
WHERE ooa.`type`=1
AND ooa.`finsh_time` >'{}'
AND ooa.`finsh_time` <'{}'
GROUP BY DATE(ooa.`finsh_time`)
) b
ON a.date=b.date
'''
scur.execute(aftersaleSql.format(first.strftime(dateFormat), today.strftime(dateFormat), first.strftime(dateFormat),
                                 today.strftime(dateFormat)))
result = scur.fetchall()
sheet.write(0, 10, '日期')
sheet.write(0, 11, '申请退货')
sheet.write(0, 12, '完成退货')
write_sheet2(result, sheet, 10)
refundsum = 0
for i, r in enumerate(result):
    if r[2] is None:
        sheet.write(i + 1, 12, 0)
    else:
        sheet.write(i + 1, 12, r[2])
        refundsum += r[2]
sheet.write(len(result)+1, 12, refundsum)

print('单品销量。。。', time.time()-stime)
sheet = wb.add_sheet('单品')
sheet.write(0, 0, '日销机型')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '数量')
sheet.write(0, 7, '月销机型')
sheet.write(0, 8, '内存')
sheet.write(0, 9, '颜色')
sheet.write(0, 10, '数量')

modelSaleSql = '''
SELECT pm.model_name,pp.`key_props` FROM panda.`odi_order` oo
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.`product_id`
LEFT JOIN panda.`pdi_model` pm
ON pm.`model_id` = pp.`model_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`pay_at` > '{}'
AND oo.`pay_at` < '{}'
'''

scur.execute(modelSaleSql.format(yesterday.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchall()
dayCount = product_count(result)
write_sheet1(dayCount, sheet, 0)

scur.execute(modelSaleSql.format(first.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchall()
monthCount = product_count(result)
write_sheet1(monthCount, sheet, 7)

print('上架机型。。。', time.time())
groundSql = '''

'''

#path = cf.get('path', 'path')
wb.save('day.xls')
scur.close()
scon.close()
print('overtime...', time.time()-stime)
