import pymysql
import config
import xlwt
import time
import decimal


class Product:
    def __init__(self, props, cycle_time):
        self.props = props
        self.cycle_time = cycle_time

start_time = time.time()
cf = config.product
src_con = pymysql.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], db=cf['db'],
                          charset=config.charset)
src_cur = src_con.cursor()
workbook = xlwt.Workbook()

props_sql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''

version_dict = config.properties_dict(src_cur, props_sql, 5)
memory_dict = config.properties_dict(src_cur, props_sql, 11)
color_dict = config.properties_dict(src_cur, props_sql, 10)

yipin_batch_sql = '''
SELECT pb.batch_no FROM panda.`pdi_batch` pb
WHERE pb.`suppiler`=272
'''
src_cur.execute(yipin_batch_sql)
result = src_cur.fetchall()
yipin = []
for r in result:
    yipin.append(str(r[0]))

print('collecting products... {}'.format(time.time()-start_time))
sale_product_sql = '''
SELECT (UNIX_TIMESTAMP(sw.`out_time`)-UNIX_TIMESTAMP(sw.`in_time`))/60/60/24 `time`,sw.`key_props` 
FROM panda.`stg_warehouse` sw
LEFT JOIN panda.`odi_order` oo
ON sw.`product_id` = oo.`product_id`
WHERE oo.`order_status` IN (1,2,4,5)
and oo.order_type in (1,2)
and sw.in_time > '2017-1-1'
and sw.out_time > '0000-00-00 00:00:00'
and sw.batch_no not in ({})
'''
src_cur.execute(sale_product_sql.format(','.join(yipin)))
print('collect finish... {}'.format(time.time()-start_time))
result = src_cur.fetchall()
sale_product_list = []
for r in result:
    if r[0] is not None and r[1] is not None:
        product = Product(r[1], r[0])
        properties = product.props.split(';')
        for feature in properties:
            f = feature.split(':')
            if f[0] == '5':
                product.version = version_dict[str(f[1])]
            if f[0] == '10':
                product.color = color_dict[str(f[1])]
            if f[0] == '11':
                product.memory = memory_dict[str(f[1])]
        sale_product_list.append(product)

sale_product_dict_count = {}
sale_product_dict_time = {}
for p in sale_product_list:
    name = p.version + ':' + p.color + ':' + p.memory
    if name in sale_product_dict_count:
        sale_product_dict_count[name] = sale_product_dict_count[name] + 1
    else:
        sale_product_dict_count[name] = 1
    if name in sale_product_dict_time:
        sale_product_dict_time[name] = sale_product_dict_time[name] + p.cycle_time
    else:
        sale_product_dict_time[name] = p.cycle_time

# print(product_dict_time)
sheet = workbook.add_sheet('sheet')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '颜色')
sheet.write(0, 2, '内存')
sheet.write(0, 3, '数量')
sheet.write(0, 4, '平均时长')
for i, p in enumerate(sale_product_dict_count):
    pv, pc, pm = p.split(':')
    sheet.write(i+1, 0, pv)
    sheet.write(i+1, 1, pc)
    sheet.write(i+1, 2, pm)
    sheet.write(i+1, 3, sale_product_dict_count[p])
    sheet.write(i+1, 4, sale_product_dict_time[p]/sale_product_dict_count[p])

store_product_sql = '''
SELECT (UNIX_TIMESTAMP(NOW())-UNIX_TIMESTAMP(sw.`in_time`))/60/60/24,sw.`key_props`,ppc.cost 
FROM panda.`stg_warehouse` sw
left join panda.pdi_product_cost ppc
on sw.product_id = ppc.product_id
WHERE sw.`warehouse_status` =1
and sw.warehouse_num not in (3,12)
'''
src_cur.execute(store_product_sql)
result = src_cur.fetchall()
store_product_list = []
for r in result:
    if r[0] is not None and r[1] is not None:
        product = Product(r[1], r[0])
        product.cost = r[2]
        properties = product.props.split(';')
        for feature in properties:
            f = feature.split(':')
            if f[0] == '5':
                product.version = version_dict[str(f[1])]
            if f[0] == '10':
                product.color = color_dict[str(f[1])]
            if f[0] == '11':
                product.memory = memory_dict[str(f[1])]
        store_product_list.append(product)

store_product_dict_count = {}
store_product_dict_time = {}
store_product_cost = {}
mis = 0
for p in store_product_list:
    name = p.version + ':' + p.color + ':' + p.memory
    if name in store_product_dict_count:
        store_product_dict_count[name] = store_product_dict_count[name] + 1
    else:
        store_product_dict_count[name] = 1
    if name in store_product_dict_time:
        store_product_dict_time[name] = store_product_dict_time[name] + p.cycle_time
    else:
        store_product_dict_time[name] = p.cycle_time
    if name in store_product_cost:
        if p.color is None:
            mis += 1
        store_product_cost[name] = store_product_cost[name] + p.cost if p.cost is not None else decimal.Decimal('2500')
    else:
        store_product_cost[name] = p.cost

print(mis)
# print(product_dict_time)
sheet.write(0, 10, '型号')
sheet.write(0, 11, '颜色')
sheet.write(0, 12, '内存')
sheet.write(0, 13, '数量')
sheet.write(0, 14, '平均在库时长')
sheet.write(0, 15, '成本')
for i, p in enumerate(store_product_dict_count):
    pv, pc, pm = p.split(':')
    sheet.write(i+1, 10, pv)
    sheet.write(i+1, 11, pc)
    sheet.write(i+1, 12, pm)
    sheet.write(i+1, 13, store_product_dict_count[p])
    sheet.write(i+1, 14, store_product_dict_time[p]/store_product_dict_count[p])
    sheet.write(i+1, 15, store_product_cost[p])


workbook.save('warehousemean.xls')
src_cur.close()
src_con.close()
