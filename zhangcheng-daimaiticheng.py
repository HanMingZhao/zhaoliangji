import config as conf
import xlwt


class Product:
    def __init__(self, prop, pid):
        self.prop = prop
        self.id = pid

propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = conf.properties_dict(conf.product_cursor, propsql, 5)
md = conf.properties_dict(conf.product_cursor, propsql, 11)
cd = conf.properties_dict(conf.product_cursor, propsql, 10)

total_amount_sql = '''
SELECT oo.`product_id`,oo.`total_amount` FROM panda.`odi_order` oo
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`product_id` NOT IN (0, 1, 84)
AND oo.`pay_at` > '2017-11-1'
'''
conf.product_cursor.execute(total_amount_sql)
result = conf.product_cursor.fetchall()
product_total_amount = {}
for r in result:
    product_total_amount[r[0]] = r[1]

cost_sql = '''
SELECT oo.`product_id`,ppc.`cost` FROM panda.`odi_order`  oo
LEFT JOIN panda.`pdi_product_cost` ppc
ON oo.`product_id` = ppc.`product_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`product_id` NOT IN (0,1,84)
AND oo.`pay_at` > '2017-11-1'
'''
conf.product_cursor.execute(cost_sql)
result = conf.product_cursor.fetchall()
product_cost = {}
for r in result:
    product_cost[r[0]] = r[1]

count_sql = '''
SELECT pp.key_props,oo.`product_id` FROM panda.`odi_order` oo
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.product_id
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`product_id` NOT IN (0,1,84)
AND oo.`pay_at` > '2017-11-1'
'''
conf.product_cursor.execute(count_sql)
result = conf.product_cursor.fetchall()
product_sku_dict = {}
for sr in result:
    if sr[0] is not None:
        p = Product(sr[0], sr[1])
        properties = p.prop.split(';')
        for f in properties:
            feature = f.split(':')
            if feature[0] == '5':
                p.version = vd[feature[1]]
            if feature[0] == '10':
                p.color = cd[feature[1]]
            if feature[0] == '11':
                p.memory = md[feature[1]]
        name = p.version + ':' + p.memory + ':' + p.color
        product_sku_dict[sr[1]] = name

count_dict = []
sum_dict = []
for pta in product_total_amount:
    earn = (product_total_amount[pta]*0.98-58-product_cost[pta])*0.5
    if product_sku_dict[pta] in count_dict:
        count_dict[product_sku_dict[pta]] = count_dict[product_sku_dict[pta]] +1
    else:
        count_dict[product_sku_dict[pta]] = 1
    if product_sku_dict[pta] in sum_dict:
        sum_dict[product_sku_dict[pta]] = sum_dict[product_sku_dict[pta]] + earn
    else:
        sum_dict[product_sku_dict[pta]] = 1

wb = xlwt.Workbook()
sheet = wb.add_sheet('sheet')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '平均提成')

for c in count_dict:
    row = len(sheet.rows)
    pv, pm, pc = c.split(':')
    sheet.write(row, 0, pv)
    sheet.write(row, 1, pm)
    sheet.write(row, 2, pc)
    sheet.write(row, 3, sum_dict[c]/count_dict[c])

wb.save('avgjunjia.xls')
conf.product_cursor.close()
conf.product_connect.close()
