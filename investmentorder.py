import pymysql as db
import config as conf
import xlwt

cf = conf.product
con = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cur = con.cursor()
wb = xlwt.Workbook()

order_sql = '''
SELECT oo.`user_id`,oo.`total_amount`,oo.`product_id`,DATE(oo.`pay_at`) FROM panda.`odi_order` oo
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`pay_at` > '2017-11-1'
AND oo.`pay_at` < '2018-1-1'
ORDER BY oo.`pay_at`
'''
cur.execute(order_sql)
order_result = cur.fetchall()
product_list = []
for r in order_result:
    product_list.append(r[2])

placeholder = '?' # For SQLite. See DBAPI paramstyle.
placeholders = ', '.join(placeholder for unused in product_list)
imei_sql = 'SELECT pp.product_id,pp.tag FROM panda.pdi_product pp WHERE pp.product_id IN (%s)' % placeholders
cur.execute(imei_sql)
imei_result = cur.fetchall()
imei_dict = {}
for ir in imei_result:
    imei_dict[ir[0]] = ir[1]

cost_sql = 'select ppc.product_id,ppc.cost from panda.pdi_product_cost ppc where ppc.product_id in (%s)' % placeholders
cur.execute(cost_sql)
cost_result = cur.fetchall()
cost_dict = {}
for cr in cost_result:
    cost_dict[cr[0]] = cr[1]

sheet = wb.add_sheet('sheet')
sheet.write(0, 0, '用户ID')
sheet.write(0, 1, '售价')
sheet.write(0, 2, '成本')
sheet.write(0, 3, '机器码')
sheet.write(0, 4, '销售时间')

for i, r in enumerate(order_result):
    sheet.write(i+1, 0, r[0])
    sheet.write(i+1, 1, r[1])
    sheet.write(i+1, 2, cost_dict[r[2]])
    sheet.write(i+1, 3, imei_dict[r[2]])
    sheet.write(i+1, 4, r[3])

wb.save(conf.path + 'orderdetail.xls')
cur.close()
con.close()
