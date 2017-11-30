import pymysql as db
import xlwt
import datetime
import configparser
import time
import os
import numpy as np


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


def sheet_head(sheety, idx):
    sheety.write(0, idx, '型号')
    sheety.write(0, idx+1, '内存')
    sheety.write(0, idx+2, '颜色')
    sheety.write(0, idx+3, '总库存')

stime = time.time()

cf = configparser.ConfigParser()
cf.read(os.path.dirname(__file__) + '/conf.conf')
option = 'db'
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

#today = datetime.datetime.strptime('2016-11-24', dateFormat)
today = datetime.datetime.today()
yesterday = today-datetime.timedelta(1)
month = today.month
year = today.year
if today.day == 1:
    month -= 1
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
sheet = wb.add_sheet('日销')
sheet_head(sheet, 0)

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

sheet = wb.add_sheet('月销')
sheet_head(sheet, 0)
scur.execute(modelSaleSql.format(first.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchall()
monthCount = product_count(result)
write_sheet1(monthCount, sheet, 0)

print('上架机型。。。', time.time())
groundSql = '''
SELECT pm.`model_name`,pp.`key_props` FROM panda.`pdi_product_track` ppt 
LEFT JOIN panda.`pdi_product` pp
ON ppt.`product_id` = pp.`product_id`
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id` = pm.`model_id`
WHERE ppt.`track_type` = 1
AND ppt.`product_status` !=3
AND ppt.`created_at` > '{}'
AND ppt.`created_at` < '{}'
GROUP BY ppt.id
'''
scur.execute(groundSql.format(yesterday.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchall()
sku = product_count(result)

sheet = wb.add_sheet('上架量')
sheet_head(sheet, 0)
write_sheet1(sku, sheet, 0)

print('预上架...', time.time()-stime)
pregroundSql = '''
SELECT pm.`model_name`,pp.`key_props` FROM panda.`stg_warehouse_switch` sws 
LEFT JOIN panda.`pdi_product` pp
ON sws.product_id = pp.product_id
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id` = pm.`model_id`
WHERE sws.`switch_status` =2
AND sws.`dst_warehouse` = 8
AND sws.`check_time`> '{}'
AND sws.`check_time`< '{}'
GROUP BY sws.`dst_warehouse`,sws.imei
'''
scur.execute(pregroundSql.format(yesterday.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchall()
presku = product_count(result)

sheet = wb.add_sheet('预上架量')
sheet_head(sheet, 0)
write_sheet1(presku, sheet, 0)

print('总库存。。。', time.time()-stime)
storeSql = '''
SELECT pm.`model_name`,sw.`key_props` FROM panda.`stg_warehouse` sw
LEFT JOIN panda.`pdi_model` pm
ON pm.`model_id` = sw.`model_id`
WHERE sw.`warehouse_status`=1
{}
'''
scur.execute(storeSql.format(''))
result = scur.fetchall()
storagesku = product_count(result)
sheet = wb.add_sheet('总库存')
sheet_head(sheet, 0)
write_sheet1(storagesku, sheet, 0)

condition = 'and sw.warehouse_num = {} '
print('上架库。。。', time.time()-stime)
scur.execute(storeSql.format(condition.format(4)))
result = scur.fetchall()
storagesku = product_count(result)
sheet = wb.add_sheet('上架库')
sheet_head(sheet, 0)
write_sheet1(storagesku, sheet, 0)

print('预上架库', time.time()-stime)
scur.execute(storeSql.format(condition.format(8)))
result = scur.fetchall()
storagesku = product_count(result)
sheet = wb.add_sheet('预上架库')
sheet_head(sheet, 0)
write_sheet1(storagesku, sheet, 0)

print('B端库...', time.time()-stime)
scur.execute(storeSql.format(condition.format(7)))
result = scur.fetchall()
storagesku = product_count(result)
sheet = wb.add_sheet('B端')
sheet_head(sheet, 0)
write_sheet1(storagesku, sheet, 0)

print('库存周转15天。。。', time.time()-stime)
print('上架库15天。。。', time.time()-stime)
roundSql = '''
SELECT pm.model_name,sw.key_props FROM panda.`stg_warehouse` sw
LEFT JOIN panda.`pdi_model` pm
ON sw.model_id = pm.model_id
WHERE sw.`warehouse_status` = 1
AND sw.warehouse_num = {}
AND (NOW()-sw.change_time)/60/60/24>15
'''
scur.execute(roundSql.format(4))
result = scur.fetchall()
roundsku = product_count(result)
sheet = wb.add_sheet('上架大于15天')
sheet_head(sheet, 0)
write_sheet1(roundsku, sheet, 0)

print('预上架库15天。。。', time.time()-stime)
scur.execute(roundSql.format(8))
result = scur.fetchall()
roundsku = product_count(result)
sheet = wb.add_sheet('预上架大于15天')
sheet_head(sheet, 0)
write_sheet1(roundsku, sheet, 0)

print('B端库15天。。。', time.time()-stime)
scur.execute(roundSql.format(7))
result = scur.fetchall()
roundsku = product_count(result)
sheet = wb.add_sheet('B端大于15天')
sheet_head(sheet, 0)
write_sheet1(roundsku, sheet, 0)

print('购买时间段。。。', time.time()-stime)
saleTimeSql = '''
SELECT HOUR(oo.pay_at),COUNT(1) FROM panda.`odi_order` oo
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`pay_at` > '{}'
AND oo.`pay_at` < '{}'
GROUP BY HOUR(oo.`pay_at`)
'''
scur.execute(saleTimeSql.format(yesterday.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchall()
sheet = wb.add_sheet('购买时间段')
sheet.write(0, 0, '时间')
sheet.write(0, 1, '数量')
daySale = np.zeros(24)
for r in result:
    daySale[int(r[0])] = r[1]
row = 1
for d in daySale:
    sheet.write(row, 0, str(row-1)+'点')
    sheet.write(row, 1, d)
    row += 1

print('pv...', time.time()-stime)
sheet = wb.add_sheet('pv')
pvSql = '''
SELECT bai.`created_at`,bai.`pv`,bai.`ip`,bai.`register` FROM panda.`boss_api_info` bai
WHERE bai.`created_at` < '{}'
AND bai.`created_at` > '{}'
ORDER BY bai.`created_at` ASC
'''
scur.execute(pvSql.format(first.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchall()
sheet.write(0, 0, '日期')
sheet.write(0, 1, 'pv')
sheet.write(0, 2, 'uv')
sheet.write(0, 3, '注册')
sheet.write(0, 4, '新增')
sheet.write(0, 5, 'ios新增')
sheet.write(0, 6, 'android新增')
sheet.write(0, 7, '日活')
sheet.write(0, 8, 'ios日活')
sheet.write(0, 9, 'android日活')
for i, r in enumerate(result):
    sheet.write(i+1, 0, r[0])
    sheet.write(i+1, 1, r[1])
    sheet.write(i+1, 2, r[2])
    sheet.write(i+1, 3, r[3])


path = cf.get('path', 'path')
wb.save('day.xls')
scur.close()
scon.close()
print('overtime...', time.time()-stime)
