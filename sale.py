import pymysql as db
import numpy as np
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
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
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
transaction = '%.2f' % (saleAmount / uv / pct * 100)

sheet.write(0, 0, '重要级')
sheet.write(0, 1, '项目')
sheet.write(0, 2, '指标')
sheet.write(0, 3, '指标明细')
sheet.write(0, 4, yesterday.strftime(dateFormat))
sheet.write_merge(1, 1, 0, 3, '星期')
sheet.write(1, 4, yesterday.strftime('%A'))
sheet.write_merge(2, 6, 0, 0, 'A')
sheet.write_merge(2, 3, 1, 1, '核心数据')
sheet.write_merge(2, 2, 2, 3, '实际销售额')
sheet.write(2, 4, saleAmount)
sheet.write_merge(3, 3, 2, 3, '实际下单数量')
sheet.write(3, 4, saleCount)
sheet.write_merge(4, 6, 1, 1, '反馈数据')
sheet.write_merge(4, 4, 2, 3, 'UV')
sheet.write(4, 4, uv)
sheet.write_merge(5, 5, 2, 3, '转化率')
sheet.write(5, 4, transaction + '%')
sheet.write_merge(6, 6, 2, 3, '客单价')
sheet.write(6, 4, pct)

path = cf.get('path', 'path')
wb.save(path + today.strftime(dateFormat) + 'pct.xls')

scur.close()
scon.close()
