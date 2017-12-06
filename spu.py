import pymysql as db
import config as conf
import time

start_time = time.time()

# cf = conf.test
cf = conf.product
src_con = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], db=cf['db'])
src_cur = src_con.cursor()

properties_sql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
version_dict = conf.properties_dict(src_cur, properties_sql, 5)

memory_dict = conf.properties_dict(src_cur, properties_sql, 11)

rate_dict = conf.properties_dict(src_cur, properties_sql, 12)

products_sql = '''
SELECT pp.`key_props` FROM panda.`pdi_product` pp
WHERE pp.`status` =1
'''
src_cur.execute(products_sql)
products_result = src_cur.fetchall()
products_set = conf.product_count(products_result, version_dict, memory_dict, rate_dict)
print(len(products_set))

src_cur.close()
src_con.close()
