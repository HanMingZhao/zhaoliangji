import redis
import config as conf
import pymysql as db
import xlwt
import time

start = time.time()
cf = conf.product
conn = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], db=cf['db'], port=cf['port'], charset=conf.char)
cur = conn.cursor()
pool = redis.ConnectionPool(host='127.0.0.1', port=36379)
r = redis.StrictRedis(connection_pool=pool, decode_responses=True)
workbook = xlwt.Workbook()
prefix = 'count_product_detail_key:'
print('2016sqlscan...', time.time()-start)
product_sql = '''
SELECT oo.product_id,pp.product_name,date(oo.pay_at) FROM panda.`odi_order` oo
left join panda.pdi_product pp 
on oo.product_id = pp.product_id
WHERE oo.order_status IN (1,2,4,5)
AND oo.order_type IN (1,2)
AND oo.pay_at > '{}'
AND oo.pay_at < '{}'
'''
cur.execute(product_sql.format('2016-1-1', '2017-1-1'))
result = cur.fetchall()
sheet = workbook.add_sheet('2016')
sheet.write(0, 0, '销售日期')
sheet.write(0, 1, '产品ID')
sheet.write(0, 2, '产品名称')
sheet.write(0, 3, '访问次数')
print('2016redisscan...', time.time()-start)
for i, res in enumerate(result):
    sheet.write(i+1, 0, str(res[2]))
    sheet.write(i+1, 1, res[0])
    sheet.write(i+1, 2, res[1])
    sheet.write(i+1, 3, str(r.get(prefix+str(res[0])), encoding='utf-8') if r.exists(prefix+str(res[0])) else 0)

print('2017sqlscan...', time.time()-start)
cur.execute(product_sql.format('2017-1-1', '2018-1-1'))
result = cur.fetchall()
sheet = workbook.add_sheet('2017')
sheet.write(0, 0, '销售日期')
sheet.write(0, 1, '产品ID')
sheet.write(0, 2, '产品名称')
sheet.write(0, 3, '访问次数')
print('2017redisscan...', time.time()-start)
for i, res in enumerate(result):
    sheet.write(i+1, 0, str(res[2]))
    sheet.write(i+1, 1, res[0])
    sheet.write(i+1, 2, res[1])
    sheet.write(i+1, 3, str(r.get(prefix+str(res[0])), encoding='utf-8') if r.exists(prefix+str(res[0])) else 0)

workbook.save('view.xls')
cur.close()
conn.close()
print('over...', time.time()-start)
