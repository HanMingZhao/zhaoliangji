import pymysql as db
import config as conf
import xlwt


class Product:
    def __init__(self, props, brand):
        self.props = props
        self.brand = brand


cf = conf.product
connect = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cursor = connect.cursor()
wb = xlwt.Workbook()

propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = conf.properties_dict(cursor, propsql, 5)
sd = conf.properties_dict(cursor, propsql, 8)
md = conf.properties_dict(cursor, propsql, 11)
cd = conf.properties_dict(cursor, propsql, 10)

grounding_sql = '''
SELECT pp.key_props,pm.`brand_name` FROM panda.`pdi_product` pp
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id` =pm.`model_id`
WHERE pp.`status` = 1
AND pp.`key_props` LIKE '%9:1;%'
'''
cursor.execute(grounding_sql)
result = cursor.fetchall()
product_list = []
for r in result:
    product = Product(r[0], r[1])
    properties = product.props.split(';')
    for f in properties:
        feature = f.split(':')
        if feature[0] == '5':
            product.version = vd[feature[1]]
        if feature[0] == '10':
            product.color = cd[feature[1]]
        if feature[0] == '11':
            product.memory = md[feature[1]]
        if feature[0] == '8':
            product.sign = sd[feature[1]]
    product_list.append(product)
product_dict = {}
for prod in product_list:
    name = prod.brand + ':' + prod.version + ':' + prod.memory + ':' + prod.color + ':' + prod.sign
    if name in product_dict:
        product_dict[name] = product_dict[name] + 1
    else:
        product_dict[name] = 1

sheet = wb.add_sheet('sheet')
sheet.write(0, 0, '品牌')
sheet.write(0, 1, '型号')
sheet.write(0, 2, '内存')
sheet.write(0, 3, '颜色')
sheet.write(0, 4, '网络制式')
sheet.write(0, 5, '数量')
for p in product_dict:
    pb, pv, pm, pc, ps = p.split(':')
    row = len(sheet.rows)
    sheet.write(row, 0, pb)
    sheet.write(row, 1, pv)
    sheet.write(row, 2, pm)
    sheet.write(row, 3, pc)
    sheet.write(row, 4, ps)
    sheet.write(row, 5, product_dict[p])

wb.save('grounding.xls')
cursor.close()
connect.close()
