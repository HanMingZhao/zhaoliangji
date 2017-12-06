# import config as conf
# import datetime as dt
# import xlwt
# import pymysql as db
# import collections
#
# cf = conf.test
# # option = cf.product
# src_con = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], db=cf['db'],
#                      charset=conf.charset)
# src_cur = src_con.cursor()
#
# wb = xlwt.Workbook()
# sheet = wb.add_sheet('sheet')
#
# properties_sql = '''
# SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
# WHERE ppv.pnid = {}
# '''
# version_dict = conf.properties_dict(src_cur, properties_sql, 5)
#
# memory_dict = conf.properties_dict(src_cur, properties_sql, 11)
#
# sale_sql = '''
# SELECT pp.`key_props` FROM panda.`odi_order` oo
# LEFT JOIN panda.`pdi_product` pp
# ON pp.`product_id` = oo.`product_id`
# LEFT JOIN panda.pdi_model pm
# on pm.model_id = pp.model_id
# WHERE oo.`order_status` IN (1,2,4,5)
# AND oo.`order_type` IN (1,2)
# AND oo.`pay_at`> '{}'
# AND oo.`pay_at`< '{}'
# ORDER BY pm.model_name
# '''
# date1 = '2017-11-20'
# date2 = '2017-11-21'
# date3 = '2017-11-22'
# date4 = '2017-11-23'
# date5 = '2017-11-24'
# date6 = '2017-11-25'
# date7 = '2017-11-26'
#
# src_cur.execute(sale_sql.format(date1, date7))
# sales_result = src_cur.fetchall()
# sales_dict = conf.product_count(sales_result, version_dict, memory_dict, None, None)
#
# model_memory_dict = collections.OrderedDict()
# for sale in sales_dict:
#     pv, pm = sale.split(':')
#     if pv in model_memory_dict:
#         model_memory_dict[pv].add(pm)
#     else:
#         memory_set = set()
#         memory_set.add(pm)
#         model_memory_dict[pv] = memory_set
#
# model_list = [m for m in model_memory_dict]
#
# sheet.write(0, 0, '型号')
# sheet.write(0, 1, '内存')
# sheet.write(0, 2, '数量')
# for i, mm in enumerate(model_memory_dict):
#     row = len(sheet.rows)
#     sheet.write(row, 0, mm)
#     for j, m in enumerate(model_memory_dict[mm]):
#         sheet.write(row+j, 1, m)
#
#
# src_cur.execute(sale_sql.format(date1, date2))
# sale_result = src_cur.fetchall()
# sale_dict = conf.product_count(sale_result, version_dict, memory_dict, None, None)
# for sale in sale_dict:
#     pv, pm = sale.split(':')
#     model_list.index(pv)
#
# # wb.save('test.xls')
#
# src_cur.close()
# src_con.close()
