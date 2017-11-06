import pymysql as db
import configparser

cf = configparser.ConfigParser()
dbhost = cf.get('db', 'db_host')
dbuser = cf.get('db', 'db_user')
dbport = cf.get('db', 'db_port')
dbpass = cf.get('db', 'db_pass')
dbase = cf.get('db', 'db_db')
src_con = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, charset='utf8')
dst_con = db.connect(host='114.215.176.190', user='root', passwd='huodao123', db='ods', port=33069, charset='utf8')

src_cur = src_con.cursor()
dst_cur = dst_con.cursor()

saleSumSql = '''SELECT DATE(oo.`create_at`),COUNT(1) FROM panda.`odi_order` oo
LEFT JOIN panda.`aci_user_info` aui
ON oo.`user_id`=aui.`user_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`create_at`>DATE(now())-1
AND oo.`create_at`<DATE(now())'''

src_cur.execute(saleSumSql)

result = src_cur.fetchone()
lastDate = result[0]
saleSum = result[1]

saleBSql = '''#销售量
SELECT DATE(oo.`create_at`),COUNT(1) FROM panda.`odi_order` oo
LEFT JOIN panda.`aci_user_info` aui
ON oo.`user_id`=aui.`user_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`create_at`>DATE(now())-1
AND oo.`create_at`<DATE(now())
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
'''
src_cur.execute(saleBSql)
result = src_cur.fetchone()
saleC = result[1]
saleB = saleSum - saleC

saleTotalSql = '''SELECT DATE(oo.`create_at`),SUM(oo.`total_amount`) FROM panda.`odi_order` oo
LEFT JOIN panda.`aci_user_info` aui
ON oo.`user_id`=aui.`user_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`create_at`>DATE(now())-1
AND oo.`create_at`<DATE(now())
'''
src_cur.execute(saleTotalSql)
result = src_cur.fetchone()
saleTotal = result[1]

saleTotalCSql = '''SELECT DATE(oo.`create_at`),SUM(oo.`total_amount`) FROM panda.`odi_order` oo
LEFT JOIN panda.`aci_user_info` aui
ON oo.`user_id`=aui.`user_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`create_at`>DATE(now())-1
AND oo.`create_at`<DATE(now())
AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享','京东店铺','机密')
'''
src_cur.execute(saleTotalCSql)
result = src_cur.fetchone()
saleTotalC = result[1]
saleTotalB = saleTotal - saleTotalC

refundSql = '''SELECT COUNT(1) FROM panda.`odi_finance_refund_record` ofrr
WHERE ofrr.re_status = 1
AND ofrr.pay_at >0
AND ofrr.pay_money>999
AND ofrr.`create_at`>DATE(now())-1
AND ofrr.`create_at`<DATE(now())
'''
src_cur.execute(refundSql)
result = src_cur.fetchone()
refund = result[0]

refundTotalSql = '''SELECT SUM(ofrr.pay_money) FROM panda.`odi_finance_refund_record` ofrr
WHERE ofrr.re_status = 1
AND ofrr.pay_at>0
AND ofrr.pay_money>999
AND ofrr.`create_at`>DATE(now())-1
AND ofrr.`create_at`<DATE(now())
'''
src_cur.execute(refundTotalSql)
result = src_cur.fetchone()
refundTotal = result[0]

paymenSql = '''select date(oo.`create_at`),oo.`payment_id`,count(1) from panda.`odi_order` oo
where oo.`order_status` in (1,2,4,5)
AND oo.`create_at`>DATE(now())-1
AND oo.`create_at`<DATE(now())
group by oo.`payment_id`,date(oo.`create_at`)
'''
src_cur.execute(paymenSql)
result = src_cur.fetchall()
aliPay = result[0][2]
wxPay = result[1][2]
BPay = result[2][2]
mixPay = result[3][2]

reviewSql = '''SELECT DATE(rr.created_at),COUNT(1) FROM panda.`rev_review` rr
WHERE rr.`created_at`>DATE(now())-1 
AND rr.`created_at`<DATE(now())
'''
src_cur.execute(reviewSql)
result = src_cur.fetchone()
review = result[1]

pvSql = '''SELECT pv,ip FROM panda.`boss_api_info` ORDER BY created_at DESC LIMIT 1
'''
src_cur.execute(pvSql)
result = src_cur.fetchone()
pv = result[0]
uv = result[1]

dst_args = (lastDate, saleSum, saleC, saleB, saleTotal, saleTotalC, saleTotalB, refund, refundTotal,
            wxPay, aliPay, mixPay, BPay, pv, uv, review)
insertSql = '''
insert into ods.ods_daily_info VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,now())
'''
dst_cur.execute(insertSql, dst_args)
dst_con.commit()

src_cur.close()
src_con.close()
dst_cur.close()
dst_con.close()
