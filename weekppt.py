import pymysql as db
import config
import collections
import xlwt
import time


class Product:
    def __init__(self, props):
        self.props = props


def product_count(sql_results, version_dict, memory_dict, color_dict, rate_dict):
    product_list = []
    for sr in sql_results:
        if sr[0] is not None:
            p = Product(sr[0])
            properties = p.props.split(';')
            for f in properties:
                feature = f.split(':')
                try:
                    if feature[0] == '5':
                        p.version = version_dict[feature[1]]
                    if feature[0] == '10':
                        p.color = color_dict[feature[1]]
                    if feature[0] == '11':
                        p.memory = memory_dict[feature[1]]
                    if feature[0] == '12':
                        p.rate = rate_dict[feature[1]]
                except:
                    print(feature)
            product_list.append(p)
    product_dict = collections.OrderedDict()
    for prod in product_list:
        if hasattr(prod, 'rate'):
            name = prod.version + ':' + prod.memory + ':' + prod.color + ':' + prod.rate
            if name in product_dict:
                product_dict[name] = product_dict[name] + 1
            else:
                product_dict[name] = 1
    return product_dict

cf = config.product
con = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=config.char)
cur = con.cursor()
wb = xlwt.Workbook()

propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = config.properties_dict(cur, propsql, 5)
md = config.properties_dict(cur, propsql, 11)
cd = config.properties_dict(cur, propsql, 10)
rd = config.properties_dict(cur, propsql, 12)

grounding_sql = '''
SELECT pp.`key_props` FROM panda.`pdi_product_track` ppt
LEFT JOIN panda.`pdi_product` pp
ON ppt.`product_id` = pp.`product_id`
WHERE ppt.`track_type`=1
AND ppt.`created_at` >'2017-12-18'
AND ppt.`created_at` <'2018-1-1'
'''
cur.execute(grounding_sql)
result = cur.fetchall()
grounding_dict = product_count(result, vd, md, cd, rd)
sheet = wb.add_sheet('上架量')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '成色')
sheet.write(0, 4, '数量')
for i, p in enumerate(grounding_dict):
    pv, pm, pc, pr = p.split(':')
    sheet.write(i+1, 0, pv)
    sheet.write(i+1, 1, pm)
    sheet.write(i+1, 2, pc)
    sheet.write(i+1, 3, pr)
    sheet.write(i+1, 4, grounding_dict[p])

sales_sql = '''
SELECT pp.key_props FROM panda.`odi_order` oo
LEFT JOIN panda.`pdi_product` pp
ON oo.`product_id` = pp.`product_id`
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`pay_at`>'2017-12-18'
AND oo.`pay_at` < '2018-1-1'
'''
cur.execute(sales_sql)
result = cur.fetchall()
sale_dict = product_count(sales_sql, vd, md, cd, rd)
sheet.write('销量')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '成色')
sheet.write(0, 4, '数量')
for i, p in enumerate(sale_dict):
    pv, pm, pc, pr = p.split(':')
    sheet.write(i+1, 0, pv)
    sheet.write(i+1, 1, pm)
    sheet.write(i+1, 2, pc)
    sheet.write(i+1, 3, pr)
    sheet.write(i+1, 4, sale_dict[p])

pv_sql = '''
SELECT bai.created_at,bai.pv,bai.ip,bai.register FROM panda.`boss_api_info` bai
WHERE bai.created_at >='2017-12-18'
AND bai.created_at < '2018-1-1'
'''
cur.execute(pv_sql)
result = cur.fetchall()
sheet = wb.add_sheet('pv')
sheet.write(0, 0, '日期')
sheet.write(0, 1, 'pv')
sheet.write(0, 2, 'uv')
sheet.write(0, 3, '注册')
for i, r in enumerate(result):
    sheet.write(i+1, 0, r[0])
    sheet.write(i+1, 1, r[1])
    sheet.write(i+1, 2, r[2])
    sheet.write(i+1, 3, r[3])

wb.save(config.path + 'kefu.xls')
cur.close()
con.close()
