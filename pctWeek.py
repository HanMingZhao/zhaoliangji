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

modelsSql = '''
SELECT pm.`model_name` FROM panda.odi_order oo
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.`product_id`
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id` = pm.`model_id`
WHERE oo.order_status IN (1,2,4,5)
AND oo.order_type IN (1,2)
AND oo.pay_at > '2017-11-17'
GROUP BY pm.`model_name`
'''
scur.execute(modelsSql)
result = scur.fetchall()

models = [r[0] for r in result]

sheet.write(0, 0, '重要级')
sheet.write(0, 1, '项目')
sheet.write(0, 2, '指标')
sheet.write(0, 3, '指标明细')
sheet.write_merge(1, 1, 0, 3, '星期', style)
sheet.write_merge(2, 6, 0, 0, 'A', style)
sheet.write_merge(2, 3, 1, 1, '核心数据', style)
sheet.write_merge(2, 2, 2, 3, '实际销售额')
sheet.write_merge(3, 3, 2, 3, '实际下单数量')
sheet.write_merge(4, 6, 1, 1, '反馈数据', style)
sheet.write_merge(4, 4, 2, 3, 'UV')
sheet.write_merge(5, 5, 2, 3, '转化率')
sheet.write_merge(6, 6, 2, 3, '客单价')
rowBottom = 6 + len(models) * 2
sheet.write_merge(7, rowBottom, 0, 0, 'B', style)
sheet.write_merge(7, rowBottom, 1, 1, '型号销售额', style)
for i, m in enumerate(models):
    x = i*2
    sheet.write_merge(x+7, x+8, 2, 2, m)
    sheet.write(x+7, 3, '销售额')
    sheet.write(x+8, 3, '销售量')

days = [today]
for i in range(8):
    days.insert(0, today-datetime.timedelta(i))

for i, day in enumerate(days):
    if day != today:
        end = day + datetime.timedelta(1)
        saleQuerySql = '''
        SELECT {} FROM panda.`odi_order` oo
        LEFT JOIN panda.`aci_user_info` aui
        ON oo.`user_id` = aui.`user_id`
        WHERE oo.`order_status` IN (1,2,4,5)
        AND oo.`order_type` IN (1,2)
        AND oo.`pay_at` > '{}'
        AND oo.`pay_at` < '{}'
        '''

        count = 'COUNT(1)'
        scur.execute(saleQuerySql.format(count, day.strftime(dateFormat), end.strftime(dateFormat)))
        result = scur.fetchone()
        saleCount = result[0]

        amount = 'SUM(oo.`total_amount`)'
        scur.execute(saleQuerySql.format(amount, day.strftime(dateFormat), end.strftime(dateFormat)))
        result = scur.fetchone()
        saleAmount = result[0]

        uvSql = '''
        select ip from panda.`boss_api_info` bai
        where bai.`created_at` = '{}'
        '''
        scur.execute(uvSql.format(day.strftime(dateFormat)))
        result = scur.fetchone()
        uv = result[0]

        pct = decimal.Decimal('%.2f' % (saleAmount / saleCount))
        transaction = '%.2f' % (saleAmount / uv / pct * 100)

        sheet.write(0, i+4, day.strftime(dateFormat))
        sheet.write(1, i+4, day.strftime('%A'))
        sheet.write(2, i+4, saleAmount)
        sheet.write(3, i+4, saleCount)
        sheet.write(4, i+4, uv)
        sheet.write(5, i+4, transaction+'%')
        sheet.write(6, i+4, pct)

        modelCountSql = '''
        SELECT pm.model_name,COUNT(1) `count`,SUM(oo.`total_amount`) `amount` FROM panda.`odi_order` oo
        LEFT JOIN panda.`pdi_product` pp
        ON oo.`product_id` = pp.product_id
        LEFT JOIN panda.`pdi_model` pm
        ON pp.model_id = pm.model_id
        LEFT JOIN panda.`aci_user_info` aui
        ON aui.user_id = oo.`user_id`
        WHERE oo.`order_status` IN (1,2,4,5)
        AND oo.`order_type` in (1,2) 
        AND oo.`pay_at` > '{}'
        AND oo.`pay_at` < '{}'
        GROUP BY pm.model_id
        '''
        scur.execute(modelCountSql.format(day.strftime(dateFormat), end.strftime(dateFormat)))
        countAndAmounts = scur.fetchall()

        for c in countAndAmounts:
            index = models.index(c[0])
            rowIndex = index*2
            sheet.write(rowIndex+7, i+4, c[2])
            sheet.write(rowIndex+8, i+4, c[1])

path = cf.get('path', 'path')
wb.save(path + today.strftime(dateFormat) + 'weekpct.xls')

scur.close()
scon.close()
