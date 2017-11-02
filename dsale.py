import pymysql as db
import xlwt
import datetime
import numpy as np

today = str(datetime.datetime.now().date())
path = '/var/www/python3/'

scon = db.connect(host='rm-bp13wnvyc2dh86ju1.mysql.rds.aliyuncs.com', user='panda_reader', passwd='zhaoliangji3503',
                  db='panda', charset='utf8')
scur = scon.cursor()

condition = '''where pm.brand_name {} like '%iphone%'
{} pm.brand_name {} like '%苹果%'
{} pm.brand_name {} like '%ipad%' 
'''

lastSales = '''SELECT pm.model_name,COUNT(1) `count` FROM 
(
SELECT oo.`product_id`,pp.model_id FROM panda.`odi_order` oo
LEFT JOIN panda.`aci_user_info` aui
ON oo.`user_id`=aui.`user_id`
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.product_id
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`create_at`>DATE(NOW())-1
AND oo.`create_at`<DATE(NOW())
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
) ooo 
LEFT JOIN panda.`pdi_model` pm
ON ooo.model_id = pm.model_id 
{}
GROUP BY pm.model_name
ORDER BY `count` {}
'''

scur.execute(lastSales.format('', 'desc'))
result = scur.fetchall()

wbTitle = '{}sale.xls'
workBook = xlwt.Workbook()
sheet1 = workBook.add_sheet('总销量')
sheet1.write(0, 0, '机型')
sheet1.write(0, 1, today + '销量')
for i, x in enumerate(result):
    sheet1.write(i+1, 0, x[0])
    sheet1.write(i+1, 1, x[1])

timesql = '''
SELECT COUNT(1) FROM panda.`odi_order` oo
LEFT JOIN panda.`aci_user_info` aui
ON oo.user_id = aui.user_id
WHERE oo.order_status IN (1,2,4,5)
AND oo.create_at > {}
AND oo.create_at < {}
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
'''
dateToday = datetime.datetime.today()
dateYesterday = dateToday - datetime.timedelta(1)
dateBeforeYesterday = dateYesterday - datetime.timedelta(1)
dateLastWeekDay = dateToday - datetime.timedelta(7)
dateBeforeLastWeekDay = dateLastWeekDay - datetime.timedelta(1)
dtformat = "%Y-%m-%d"
scur.execute(timesql.format(dateYesterday.strftime(dtformat), dateToday.strftime(dtformat)))
lastCount = scur.fetchone()[0]
scur.execute(timesql.format(dateBeforeYesterday.strftime(dtformat), dateYesterday.strftime(dtformat)))
beforeLast = scur.fetchone()[0]
scur.execute(timesql.format(dateBeforeLastWeekDay.strftime(dtformat), dateLastWeekDay.strftime(dtformat)))
lastWeek = scur.fetchone()[0]

sheet1.write(0, 5, '同比')
sheet1.write(0, 6, '环比')
sheet1.write(1, 5, '%.2f%%' % ((lastCount - lastWeek)/lastWeek * 100))
sheet1.write(1, 6, '%.2f%%' % ((lastCount - beforeLast)/beforeLast * 100))

scur.execute(lastSales.format(condition.format('', 'or', '', 'or', ''), 'asc'))
result = scur.fetchall()
sheet2 = workBook.add_sheet('苹果龙虎榜')
sheet2.write(0, 0, '机型')
sheet2.write(0, 1, '苹果销量龙虎榜')
for i, x in enumerate(result):
    sheet2.write(i+1, 0, x[0])
    sheet2.write(i+1, 1, x[1])

scur.execute(lastSales.format(condition.format('not', 'and', 'not', 'and', 'not'), 'asc'))
result = scur.fetchall()
sheet3 = workBook.add_sheet("安卓龙虎榜")
sheet3.write(0, 0, '机型')
sheet3.write(0, 1, '安卓销量龙虎榜')
for i, x in enumerate(result):
    sheet3.write(i+1, 0, x[0])
    sheet3.write(i+1, 1, x[1])

grounding = '''
SELECT p.model_name,COUNT(1) `count` FROM (
SELECT MIN(ppt.`product_id`),pm.`model_name`,pp.* FROM panda.`pdi_product_track` ppt 
LEFT JOIN panda.`pdi_product` pp
ON ppt.`product_id` = pp.`product_id`
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id` = pm.`model_id`
WHERE ppt.`track_type` = 1
AND ppt.`created_at` > DATE(NOW())-1
AND ppt.`created_at` < DATE(NOW())
GROUP BY ppt.id
)p
GROUP BY p.model_name
ORDER BY `count` DESC
'''

modelSet = set()
scur.execute(grounding)
gresult = scur.fetchall()
for r in gresult:
    modelSet.add(r[0])

pregrounding = '''
SELECT sp.model_name,COUNT(1) `count` FROM (
SELECT MIN(sws.w_switch_id),pm.`model_name` FROM panda.`stg_warehouse_switch` sws 
LEFT JOIN panda.`pdi_product` pp
ON sws.product_id = pp.product_id
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id` = pm.`model_id`
WHERE sws.`switch_status` =2
AND sws.`dst_warehouse` = 8
AND sws.`check_time`> DATE(NOW())-1
AND sws.`check_time`< DATE(NOW())
GROUP BY sws.`dst_warehouse`,sws.imei
) sp
GROUP BY sp.model_name
ORDER BY `count` DESC
'''
scur.execute(pregrounding)
pregresult = scur.fetchall()
for r in pregresult:
    modelSet.add(r[0])

groundings = np.zeros((len(modelSet), 3)).tolist()
for s, n in zip(modelSet, groundings):
    n[0] = s

for r in gresult:
    for n in groundings:
        if r[0] == n[0]:
            n[1] = r[1]

for r in pregresult:
    for n in groundings:
        if r[0] == n[0]:
            n[2] = r[1]

sheet4 = workBook.add_sheet('上架统计')
sheet4.write(0, 0, '机型')
sheet4.write(0, 1, '上架量')
sheet4.write(0, 2, '预上架量')
for i, n in enumerate(groundings):
    sheet4.write(i+1, 0, n[0])
    sheet4.write(i+1, 1, n[1])
    sheet4.write(i+1, 2, n[2])

storage = '''
SELECT pm.model_name,COUNT(1) `count` FROM panda.`stg_warehouse` sw
LEFT JOIN panda.`pdi_model` pm
ON sw.model_id = pm.`model_id`
WHERE sw.`warehouse_status` = 1
GROUP BY pm.model_name
ORDER BY `count` DESC
'''
scur.execute(storage)
result = scur.fetchall()
sheet5 = workBook.add_sheet('库存统计')
sheet5.write(0, 0, '机型')
sheet5.write(0, 1, '库存数量')
sheet5.write(0, 2, '库存占比')
storesum = 0
for r in result:
    storesum += r[1]

for i, r in enumerate(result):
    sheet5.write(i+1, 0, r[0])
    sheet5.write(i+1, 1, r[1])
    sheet5.write(i+1, 2, '%.2f%%' % (r[1]/storesum*100))

sku = []
modelsql = '''
SELECT DISTINCT(pm.model_name) FROM
(
SELECT sw.*,pp.`buy_at` FROM panda.`stg_warehouse` sw
LEFT JOIN panda.`pdi_product` pp ON sw.`product_id` = pp.`product_id`
) psw
LEFT JOIN panda.`pdi_model`  pm  ON  psw.model_id = pm.model_id
WHERE psw.warehouse_status = 1
ORDER BY pm.model_name 
'''
scur.execute(modelsql)
models = scur.fetchall()
for m in models:
    sku.append(m[0])

oneday = np.zeros(len(sku), dtype=int)
threeday = np.zeros(len(sku), dtype=int)
sevenday = np.zeros(len(sku), dtype=int)
fifteenday = np.zeros(len(sku), dtype=int)
thirtyday = np.zeros(len(sku), dtype=int)
outthirtyday = np.zeros(len(sku), dtype=int)

storesql = '''
SELECT bt.model_name,COUNT(1) FROM 
    (
    SELECT pm.model_name,psw.product_id,psw.product_name,psw.imei,
    IF(psw.change_time>0,(UNIX_TIMESTAMP(NOW())-UNIX_TIMESTAMP(psw.change_time))/60/60/24,
    (UNIX_TIMESTAMP(NOW())-UNIX_TIMESTAMP(psw.in_time))/60/60/24) `times` FROM
    (
    SELECT sw.*,pp.`buy_at` FROM panda.`stg_warehouse` sw
    LEFT JOIN panda.`pdi_product` pp ON sw.`product_id` = pp.`product_id`
    ) psw
    LEFT JOIN panda.`pdi_model`  pm  ON  psw.model_id = pm.model_id
    WHERE psw.warehouse_status = 1
    AND psw.warehouse_num in (2,4,8)
    ORDER BY pm.model_name ,times,imei
    )bt
    where {}
    GROUP BY bt.model_name
'''
scur.execute(storesql.format('bt.times < 1 '))
results = scur.fetchall()
for r in results:
    oneday[sku.index(r[0])] = r[1]

scur.execute(storesql.format('bt.times > 1 and bt.times <3 '))
results = scur.fetchall()
for r in results:
    threeday[sku.index(r[0])] = r[1]

scur.execute(storesql.format('bt.times > 3 and  bt.times < 7'))
results = scur.fetchall()
for r in results:
    sevenday[sku.index(r[0])] = r[1]

scur.execute(storesql.format('bt.times > 7 and bt.times < 15 '))
results = scur.fetchall()
for r in results:
    fifteenday[sku.index(r[0])] = r[1]

scur.execute(storesql.format('bt.times > 15 and bt.times < 30 '))
results = scur.fetchall()
for r in results:
    thirtyday[sku.index(r[0])] = r[1]

scur.execute(storesql.format('bt.times > 30 '))
results = scur.fetchall()
for r in results:
    outthirtyday[sku.index(r[0])] = r[1]

sku.insert(0, '周期')
sku.append('库存周转占比')
l1 = [int(x) for x in oneday]
l1.append(sum(l1))
l3 = [int(x) for x in threeday]
l3.append(sum(l3))
l7 = [int(x) for x in sevenday]
l7.append(sum(l7))
l15 = [int(x) for x in fifteenday]
l15.append(sum(l15))
l30 = [int(x) for x in thirtyday]
l30.append(sum(l30))
g30 = [int(x) for x in outthirtyday]
g30.append(sum(g30))
l1.insert(0, '小于1天内')
l3.insert(0, '小于3天内')
l7.insert(0, '小于7天内')
l15.insert(0, '小于15天内')
l30.insert(0, '小于30天内')
g30.insert(0, '大于30天')

matrix = [sku, l1, l3, l7, l15, l30, g30]
matrix2 = np.matrix(matrix)
matrix3 = matrix2.transpose()
matrix4 = matrix3.tolist()
sheet6 = workBook.add_sheet('周转统计')
sheet6.write(0, 0, 'sku')
sheet6.write(0, 1, '小于1天内')
sheet6.write(0, 2, '小于3天内')
sheet6.write(0, 3, '小于7天内')
sheet6.write(0, 4, '小于15天内')
sheet6.write(0, 5, '小于30天内')
sheet6.write(0, 6, '大于30天')

matrix4.reverse()
for i, r in enumerate(matrix4[0]):
    sheet6.write(1, i, r)

workBook.save(path + wbTitle.format(today))

scur.close()
scon.close()
