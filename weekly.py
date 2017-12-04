import pymysql as db
import datetime as dt
import xlwt
import numpy as np
import configparser

cf = configparser.ConfigParser()
cf.read('conf.conf')
dbhost = cf.get('db', 'host')
dbuser = cf.get('db', 'user')
dbport = cf.getint('db', 'port')
dbpass = cf.get('db', 'pass')
dbase = cf.get('db', 'db')
scon = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, charset='utf8')
scur = scon.cursor()
workBook = xlwt.Workbook()
today = str(dt.datetime.today().date())
path = cf.get('path', 'path')
dateToday = dt.datetime.today()
dateLastWeekDay = dateToday - dt.timedelta(7)
dtformat = "%Y-%m-%d"

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

weekSaleSql = '''
SELECT DATE(oo.`create_at`),COUNT(1) FROM panda.`odi_order` oo
LEFT JOIN panda.`aci_user_info` aui
ON oo.`user_id` = aui.`user_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`create_at` > '{}'
AND oo.`create_at` < '{}'
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
GROUP BY DATE(oo.`create_at`)
'''
scur.execute(weekSaleSql.format(dateLastWeekDay.strftime(dtformat), dateToday.strftime(dtformat)))
weekSales = scur.fetchall()
sheet1 = workBook.add_sheet('周销量')
sheet1.write(0, 0, '日期')
sheet1.write(0, 1, '每日销售总量')
for i, sale in enumerate(weekSales):
    sheet1.write(i+1, 0, sale[0])
    sheet1.write(i+1, 1, sale[1])
sheet1.write(0, 4, len(weekSales))

versionSql = '''
SELECT pm.`model_name`,COUNT(1) `count` FROM panda.`odi_order` oo
LEFT JOIN panda.`aci_user_info` aui
ON oo.`user_id` = aui.`user_id`
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.`product_id`
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id` = pm.model_id
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`create_at` > '{}'
AND oo.`create_at` < '{}'
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
GROUP BY pm.`model_name`
ORDER BY `count` DESC
'''
scur.execute(versionSql.format(dateLastWeekDay.strftime(dtformat), dateToday.strftime(dtformat)))
versionSales = scur.fetchall()
sheet2 = workBook.add_sheet('销售占比')
sheet2.write(0, 0, '机型')
sheet2.write(0, 1, '销量')
for i, version in enumerate(versionSales):
    sheet2.write(i+1, 0, version[0])
    sheet2.write(i+1, 1, version[1])

featureSql = '''
SELECT pp.`key_props` FROM panda.`odi_order` oo
LEFT JOIN panda.`aci_user_info` aui
ON oo.`user_id` = aui.`user_id`
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.`product_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`create_at` > '{}'
AND oo.`create_at` < '{}'
{}
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
'''
scur.execute(featureSql.format(dateLastWeekDay.strftime(dtformat), dateToday.strftime(dtformat), ''))
properties = scur.fetchall()
colorSales = {}
memorySales = {}
for prop in properties:
    p = prop[0].split(';')
    for feature in p:
        f = feature.split(':')
        if f[0] == '10':
            if cd[f[1]] in colorSales:
                colorSales[cd[f[1]]] = colorSales[cd[f[1]]] + 1
            else:
                colorSales[cd[f[1]]] = 1
        if f[0] == '11':
            if md[f[1]] in memorySales:
                memorySales[md[f[1]]] = memorySales[md[f[1]]] + 1
            else:
                memorySales[md[f[1]]] = 1
        continue

sheet2.write(0, 3, '颜色')
sheet2.write(0, 4, '销量')
for i, color in enumerate(colorSales):
    sheet2.write(i+1, 3, color)
    sheet2.write(i+1, 4, colorSales[color])

sheet2.write(0, 6, '内存')
sheet2.write(0, 7, '销量')
for i, memory in enumerate(memorySales):
    sheet2.write(i+1, 6, memory)
    sheet2.write(i+1, 7, memorySales[memory])

saleCountSql = '''SELECT pm.model_id,pm.model_name,COUNT(1) `count` FROM 
(
SELECT oo.`product_id`,pp.model_id FROM panda.`odi_order` oo
LEFT JOIN panda.`aci_user_info` aui
ON oo.`user_id`=aui.`user_id`
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.product_id
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`create_at`>'{}'
AND oo.`create_at`<'{}'
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
) ooo 
LEFT JOIN panda.`pdi_model` pm
ON ooo.model_id = pm.model_id 
{}
GROUP BY pm.model_name
ORDER BY `count` {}
'''
condition = '''where pm.brand_name {} like '%iphone%'
{} pm.brand_name {} like '%苹果%'
{} pm.brand_name {} like '%ipad%' 
'''

scur.execute(saleCountSql.format(dateLastWeekDay.strftime(dtformat), dateToday.strftime(dtformat),
                                 condition.format('', 'or', '', 'or', ''), 'asc'))
result = scur.fetchall()
sheet3 = workBook.add_sheet('苹果龙虎榜')
sheet3.write(0, 0, '机型')
sheet3.write(0, 1, '苹果销量龙虎榜')
for i, x in enumerate(result):
    sheet3.write(i+1, 0, x[1])
    sheet3.write(i+1, 1, x[2])

scur.execute(saleCountSql.format(dateLastWeekDay.strftime(dtformat), dateToday.strftime(dtformat),
                                 condition.format('not', 'and', 'not', 'and', 'not'), 'asc'))
result = scur.fetchall()
sheet4 = workBook.add_sheet("安卓龙虎榜")
sheet4.write(0, 0, '机型')
sheet4.write(0, 1, '安卓销量龙虎榜')
for i, x in enumerate(result):
    sheet4.write(i+1, 0, x[1])
    sheet4.write(i+1, 1, x[2])

models = []
scur.execute(saleCountSql.format(dateLastWeekDay.strftime(dtformat), dateToday.strftime(dtformat),
                                 condition.format('', 'or', '', 'or', ''), 'desc limit 5'))
result = scur.fetchall()
for model in result:
    models.append(model[0])
scur.execute(saleCountSql.format(dateLastWeekDay.strftime(dtformat), dateToday.strftime(dtformat),
                                 condition.format('not', 'and', 'not', 'and', 'not'), 'desc limit 5'))
result = scur.fetchall()
for model in result:
    models.append(model[0])
for model in models:
    modelCondition = 'AND pp.model_id = {} '
    scur.execute(featureSql.format(dateLastWeekDay.strftime(dtformat), dateToday.strftime(dtformat),
                                   modelCondition.format(model)))
    modelProperties = scur.fetchall()
    colorCount = {}
    memoryCount = {}
    for modelProp in modelProperties:
        prop = modelProp[0].split(';')
        for feature in prop:
            f = feature.split(':')
            if f[0] == '10':
                if cd[f[1]] in colorCount:
                    colorCount[cd[f[1]]] = colorCount[cd[f[1]]] + 1
                else:
                    colorCount[cd[f[1]]] = 1
            if f[0] == '11':
                if md[f[1]] in memoryCount:
                    memoryCount[md[f[1]]] = memoryCount[md[f[1]]] + 1
                else:
                    memoryCount[md[f[1]]] = 1
            continue
    sheet = workBook.add_sheet(vd[str(model)])
    sheet.write(0, 0, '颜色')
    sheet.write(0, 1, '数量')
    for i, x in enumerate(colorCount):
        sheet.write(i+1, 0, x)
        sheet.write(i+1, 1, colorCount[x])
    sheet.write(0, 3, '内存')
    sheet.write(0, 4, '数量')
    for i, x in enumerate(memoryCount):
        sheet.write(i+1, 3, x)
        sheet.write(i+1, 4, memoryCount[x])

groundingSql = '''
SELECT DATE(p.created_at) `time`,COUNT(1) `count` FROM (
SELECT DISTINCT(ppt.id),ppt.`created_at` FROM panda.`pdi_product_track` ppt 
LEFT JOIN panda.`pdi_product` pp
ON ppt.`product_id` = pp.`product_id`
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id` = pm.`model_id`
WHERE ppt.`track_type` = 1
AND ppt.`created_at` > '{}'
AND ppt.`created_at` < '{}'
GROUP BY ppt.id
)p
GROUP BY `time`
'''
scur.execute(groundingSql.format(dateLastWeekDay.strftime(dtformat), dateToday.strftime(dtformat)))
result = scur.fetchall()
sheet5 = workBook.add_sheet('上架')
sheet5.write(0, 0, '日期')
sheet5.write(0, 1, '上架')
sheet5.write(0, 2, '销量')
for i, g in enumerate(zip(result, weekSales)):
    print(i, g[0][1], g[1][1])
    sheet5.write(i+1, 0, g[0][0])
    sheet5.write(i+1, 1, g[0][1])
    sheet5.write(i+1, 2, g[1][1])

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

workBook.save(path + today + 'week.xls')

scur.close()
scon.close()
