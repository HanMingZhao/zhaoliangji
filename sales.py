import config as conf
import pymysql as db
import xlwt
import time
import collections


class Product:
    def __init__(self, props):
        self.props = props


def product_count(result, vd, md, cd, ed, rd):
    product_list = []
    for r in result:
        if r[0] is not None:
            p = Product(r[0])
            properties = p.props.split(';')
            for feature in properties:
                f = feature.split(':')
                if f[0] == '5':
                    p.version = vd[f[1]]
                if f[0] == '9':
                    p.edition = ed[f[1]]
                if f[0] == '10':
                    p.color = cd[f[1]]
                if f[0] == '11':
                    p.memory = md[f[1]]
                if f[0] == '12':
                    p.rate = rd[f[1]]
            if hasattr(p, 'rate'):
                product_list.append(p)
    product_dict = collections.OrderedDict()
    for prod in product_list:
        name = prod.version + ':' + prod.memory + ':' + prod.color + ':' + prod.edition + ':' + prod.rate
        if name in product_dict:
            product_dict[name] = product_dict[name] + 1
        else:
            product_dict[name] = 1
    return product_dict

cf = conf.product
con = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], db=cf['db'], charset=conf.char)
cur = con.cursor()
workbook = xlwt.Workbook()
start_time = time.time()
print('属性扫描...', time.time()-start_time)
props_sql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
version_dict = conf.properties_dict(cur, props_sql, 5)
memory_dict = conf.properties_dict(cur, props_sql, 11)
color_dict = conf.properties_dict(cur, props_sql, 10)
edition_dict = conf.properties_dict(cur, props_sql, 9)
rate_dict = conf.properties_dict(cur, props_sql, 12)
print('销售加载...', time.time()-start_time)
sales_sql = '''
select pp.`key_props` from panda.`odi_order` oo
left join panda.`pdi_product` pp
on oo.product_id = pp.`product_id`
where oo.order_status in (1,2,4,5)
and oo.order_type in (1,2)
and oo.pay_at > '2017-10-1'
'''
cur.execute(sales_sql)
result = cur.fetchall()
sale_dict = product_count(result, version_dict, memory_dict, color_dict, edition_dict, rate_dict)
sheet = workbook.add_sheet('sheet')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '版本')
sheet.write(0, 4, '成色')
sheet.write(0, 5, '数量')
for i, s in enumerate(sale_dict):
    pv, pm, pc, pe, pr = s.split(':')
    sheet.write(i+1, 0, pv)
    sheet.write(i+1, 1, pm)
    sheet.write(i+1, 2, pc)
    sheet.write(i+1, 3, pe)
    sheet.write(i+1, 4, pr)
    sheet.write(i+1, 5, sale_dict[s])

workbook.save(conf.path + 'edition.xls')
cur.close()
con.close()
