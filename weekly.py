import pymysql as db
import datetime as dt
import xlwt
import numpy as np

scon = db.connect(host='rm-bp13wnvyc2dh86ju1.mysql.rds.aliyuncs.com', user='panda_reader', passwd='zhaoliangji3503',
                  db='panda', charset='utf8')
scur = scon.cursor()
workBook = xlwt.Workbook()
today = str(dt.datetime.today().date())

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
AND oo.`create_at` > DATE(NOW())-7
AND oo.`create_at` < DATE(NOW())
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
GROUP BY DATE(oo.`create_at`)
'''
scur.execute(weekSaleSql)
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
AND oo.`create_at` > DATE(NOW())-7
AND oo.`create_at` < DATE(NOW())
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
GROUP BY pm.`model_name`
ORDER BY `count` DESC
'''
scur.execute(versionSql)
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
AND oo.`create_at` > DATE(NOW())-7
AND oo.`create_at` < DATE(NOW())
{}
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
'''
scur.execute(featureSql.format(''))
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
AND oo.`create_at`>DATE(NOW())-7
AND oo.`create_at`<DATE(NOW())
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

scur.execute(saleCountSql.format(condition.format('', 'or', '', 'or', ''), 'asc'))
result = scur.fetchall()
sheet3 = workBook.add_sheet('苹果龙虎榜')
sheet3.write(0, 0, '机型')
sheet3.write(0, 1, '苹果销量龙虎榜')
for i, x in enumerate(result):
    sheet3.write(i+1, 0, x[1])
    sheet3.write(i+1, 1, x[2])

scur.execute(saleCountSql.format(condition.format('not', 'and', 'not', 'and', 'not'), 'asc'))
result = scur.fetchall()
sheet4 = workBook.add_sheet("安卓龙虎榜")
sheet4.write(0, 0, '机型')
sheet4.write(0, 1, '安卓销量龙虎榜')
for i, x in enumerate(result):
    sheet4.write(i+1, 0, x[1])
    sheet4.write(i+1, 1, x[2])

models = []
scur.execute(saleCountSql.format(condition.format('', 'or', '', 'or', ''), 'desc limit 5'))
result = scur.fetchall()
for model in result:
    models.append(model[0])
scur.execute(saleCountSql.format(condition.format('not', 'and', 'not', 'and', 'not'), 'desc limit 5'))
result = scur.fetchall()
for model in result:
    models.append(model[0])
for model in models:
    modelCondition = 'AND pp.model_id = {} '
    scur.execute(featureSql.format(modelCondition.format(model)))
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

workBook.save(today + 'week.xls')

scur.close()
scon.close()
