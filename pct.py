import pymysql as db
import time
import datetime
import xlwt
import configparser
import decimal

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
sheet = wb.add_sheet('sheet')
alignment = xlwt.Alignment()
alignment.horz = alignment.HORZ_CENTER
alignment.vert = alignment.VERT_CENTER
style = xlwt.XFStyle()
style.alignment = alignment

startTime = time.time()

today = datetime.datetime.today()
dateFormat = '%Y-%m-%d'
tomorrow = today + datetime.timedelta(1)
yesterday = today - datetime.timedelta(1)

saleQuerySql = '''
select {} from panda.`odi_order` oo
left join panda.`aci_user_info` aui
on oo.`user_id` = aui.`user_id`
where oo.`order_status` in (1,2,4,5)
and oo.`pay_at` > '{}'
and oo.`pay_at` < '{}'
AND oo.order_type in (1,2)
'''
count = 'COUNT(1)'
scur.execute(saleQuerySql.format(count, yesterday.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchone()
saleCount = result[0]

amount = 'SUM(oo.`total_amount`)'
scur.execute(saleQuerySql.format(amount, yesterday.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchone()
saleAmount = result[0]

uvSql = '''
select ip from panda.`boss_api_info` bai
where bai.`created_at` = '{}'
'''
scur.execute(uvSql.format(yesterday.strftime(dateFormat)))
result = scur.fetchone()
uv = result[0]

pct = decimal.Decimal('%.2f' % (saleAmount / saleCount))
transaction = '%.2f' % (saleAmount / uv / pct * 100 if uv > 0 else 0)

sheet.write(0, 0, '重要级')
sheet.write(0, 1, '项目')
sheet.write(0, 2, '指标')
sheet.write(0, 3, '指标明细')
sheet.write(0, 4, yesterday.strftime(dateFormat))
sheet.write_merge(1, 1, 0, 3, '星期', style)
sheet.write(1, 4, yesterday.strftime('%A'))
sheet.write_merge(2, 6, 0, 0, 'A', style)
sheet.write_merge(2, 3, 1, 1, '核心数据', style)
sheet.write_merge(2, 2, 2, 3, '实际销售额')
sheet.write(2, 4, saleAmount)
sheet.write_merge(3, 3, 2, 3, '实际下单数量')
sheet.write(3, 4, saleCount)
sheet.write_merge(4, 6, 1, 1, '反馈数据', style)
sheet.write_merge(4, 4, 2, 3, 'UV')
sheet.write(4, 4, uv)
sheet.write_merge(5, 5, 2, 3, '转化率')
sheet.write(5, 4, transaction + '%')
sheet.write_merge(6, 6, 2, 3, '客单价')
sheet.write(6, 4, pct)

modelCountSql = '''
SELECT pm.model_name,COUNT(1) `count`,SUM(oo.`total_amount`) `amount` FROM panda.`odi_order` oo
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.product_id
LEFT JOIN panda.`pdi_model` pm
ON pp.model_id = pm.model_id
LEFT JOIN panda.`aci_user_info` aui
ON aui.user_id = oo.`user_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`pay_at` > '{}'
AND oo.`pay_at` < '{}'
AND oo.order_type in (1,2)
GROUP BY pm.model_name
ORDER BY `count` DESC
'''
modelCount = scur.execute(modelCountSql.format(yesterday.strftime(dateFormat), today.strftime(dateFormat)))
result = scur.fetchall()
rowBottom = 6 + modelCount * 2
sheet.write_merge(7, rowBottom, 0, 0, 'B', style)
sheet.write_merge(7, rowBottom, 1, 1, '型号销售额', style)
for i, r in enumerate(result):
    x = i*2
    sheet.write_merge(x+7, x+8, 2, 2, r[0])
    sheet.write(x+7, 3, '销售额')
    sheet.write(x+8, 3, '销售量')
    sheet.write(x+7, 4, r[2])
    sheet.write(x+8, 4, r[1])

path = cf.get('path', 'path')
wb.save(path + today.strftime(dateFormat) + 'pct.xls')

scur.close()
scon.close()
