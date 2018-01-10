import pymysql as db
import config as conf
import xlwt
import numpy as np
import pandas as pd


class Product:
    def __init__(self, props, price):
        self.props = props
        self.price = price


def product_count(sql_results, version_dict, memory_dict):
    product_list = []
    for sr in sql_results:
        if sr[0] is not None:
            p = Product(sr[0], sr[1])
            properties = p.props.split(';')
            for f in properties:
                feature = f.split(':')
                if feature[0] == '5':
                    p.version = version_dict[feature[1]]
                if feature[0] == '11':
                    p.memory = memory_dict[feature[1]]
            product_list.append(p)
    return product_list

cf = conf.new_test
connect = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cursor = connect.cursor()
wb = xlwt.Workbook()

propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = conf.properties_dict(cursor, propsql, 5)
md = conf.properties_dict(cursor, propsql, 11)
cd = conf.properties_dict(cursor, propsql, 10)

props_sql = '''
SELECT pp.`key_props`,pp.price FROM panda.`pdi_product` pp
WHERE pp.`status` = 1
AND pp.`key_props` LIKE '%12:26;%'
'''
cursor.execute(props_sql)
result = cursor.fetchall()
grounding_list = product_count(result, vd, md)
series_list = []
for r in grounding_list:
    series_list.append((r.version, r.memory, r.price))
frame = pd.DataFrame(series_list, columns=['型号', '内存', '价格'])
means = frame['价格'].groupby([frame['型号'], frame['内存']]).mean()


# sheet = wb.add_sheet('sheet')
# sheet.write(0, 0, '型号')
# sheet.write(0, 1, '内存')
# sheet.write(0, 2, '数量')
# sheet.write(0, 3, '价格')
# for r in grounding_list:
#     row = len(sheet.rows)
#     sheet.write(row, 0, r.version)
#     sheet.write(row, 1, r.memory)
#     sheet.write(row, 2, 1)
#     sheet.write(row, 3, r.price)

# wb.save('95avg.xls')
cursor.close()
connect.close()
