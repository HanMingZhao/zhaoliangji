import pymysql as db
import xlwt
from datetime import datetime
import numpy as np

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

wbTitle = '{}销售.xls'
workBook = xlwt.Workbook()
sheet1 = workBook.add_sheet('总销量')
sheet1.write(0, 0, '机型')
sheet1.write(0, 1, '数量')
for i, x in enumerate(result):
    sheet1.write(i+1, 0, x[0])
    sheet1.write(i+1, 1, x[1])

scur.execute(lastSales.format(condition.format('', 'or', '', 'or', ''), 'asc'))
result = scur.fetchall()
sheet2 = workBook.add_sheet('苹果龙虎榜')
sheet2.write(0, 0, '机型')
sheet2.write(0, 1, '数量')
for i, x in enumerate(result):
    sheet2.write(i+1, 0, x[0])
    sheet2.write(i+1, 1, x[1])

scur.execute(lastSales.format(condition.format('not', 'and', 'not', 'and', 'not'), 'asc'))
result = scur.fetchall()
sheet3 = workBook.add_sheet("安卓龙虎榜")
sheet3.write(0, 0, '机型')
sheet3.write(0, 1, '数量')
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
for i, r in enumerate(result):
    sheet5.write(i+1, 0, r[0])
    sheet5.write(i+1, 1, r[1])

workBook.save(wbTitle.format(str(datetime.now().date())))

scur.close()
scon.close()
