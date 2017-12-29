# -*- coding:utf-8 -*-
import datetime as dt
import collections

product = {'host': 'rm-m5etsh5q078zz937i.mysql.rds.aliyuncs.com',
           'user': 'zlj_reader',
           'pass': 'h=DGhEXKRq38gTtH',
           'port': 3306,
           'db': 'panda'}

test = {'host': '114.215.176.190',
        'user': 'root',
        'pass': 'huodao123',
        'port': 33069,
        'db': 'panda'}

new_test = {'host': 'rm-bp13wnvyc2dh86ju1o.mysql.rds.aliyuncs.com',
            'user': 'jaxmysql',
            'pass': 'wjnuEf0dns6PEAX1',
            'port': 3306,
            'db': 'panda'}

target = 5500

path = '/var/www/python/'

char = 'utf8'

date_format = '%Y-%m-%d'

today = dt.datetime.today()
yesterday = today - dt.timedelta(1)
last_week_day = today - dt.timedelta(7)


class Product:
    def __init__(self, props):
        self.props = props


def properties_dict(db_cursor, sql, prop_id):
    dic = {}
    db_cursor.execute(sql.format(prop_id))
    sql_result = db_cursor.fetchall()
    for s in sql_result:
        dic[str(s[0])] = s[1]
    return dic


def product_count(sql_results, version_dict, memory_dict, color_dict):
    product_list = []
    for sr in sql_results:
        if sr[0] is not None:
            p = Product(sr[0])
            properties = p.props.split(';')
            for f in properties:
                feature = f.split(':')
                if feature[0] == '5':
                    p.version = version_dict[feature[1]]
                if feature[0] == '10':
                    p.color = color_dict[feature[1]]
                if feature[0] == '11':
                    p.memory = memory_dict[feature[1]]
            product_list.append(p)
    product_dict = collections.OrderedDict()
    for prod in product_list:
        name = prod.version + ':' + prod.memory + ':' + prod.color
        if name in product_dict:
            product_dict[name] = product_dict[name] + 1
        else:
            product_dict[name] = 1
    return product_dict


def write_sheet1(count_dict, sheety, idx):
    for n, obj in enumerate(count_dict):
        objv, objm, objc = obj.split(':')
        sheety.write(n+1, idx, objv)
        sheety.write(n+1, idx+1, objm)
        sheety.write(n+1, idx+2, objc)
        sheety.write(n+1, idx+3, count_dict[obj])


def sheet_head(sheety, idx):
    sheety.write(0, idx, '型号')
    sheety.write(0, idx+1, '内存')
    sheety.write(0, idx+2, '颜色')
    sheety.write(0, idx+3, '总库存')


version_memory_dict = {4: {21, 23, 24}, 5: {21, 22, 23}, 14: {21, 22, 23, 24}, 15: {21, 22, 23, 24}, 16: {21, 23, 24},
                       114: {21, 22, 23, 24}, 127: {22, 24, 128}, 129: {22, 24, 128}, 525: {23, 128}, 544: {23, 128},
                       618: {23, 128}}

version_color_dict = {4: {17, 19, 20}, 5: {17, 19, 20}, 14: {17, 18, 19, 20}, 15: {17, 18, 19, 20}, 16: {17, 19, 20},
                      114: {17, 18, 19, 20}, 127: {17, 18, 19, 20, 133, 401}, 129: {17, 18, 19, 20, 133, 401},
                      525: {440, 17, 19}, 544: {17, 19, 440}, 618: {19, 440}}
