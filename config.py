# -*- coding:utf-8 -*-
import datetime as dt

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

path = '/var/www/python/'

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


def product_count(sql_results, version_dict, memory_dict, color_dict, rate_dict):
    product_list = []
    for sr in sql_results:
        if sr[0] is not None:
            p = Product(sr[0])
            properties = p.props.split(';')
            if '12:' in p.props:
                for f in properties:
                    feature = f.split(':')
                    if feature[0] == '5':
                        p.version = version_dict[feature[1]]
                    if feature[0] == '10':
                        p.color = color_dict[feature[1]]
                    if feature[0] == '11':
                        p.memory = memory_dict[feature[1]]
                    if feature[0] == '12':
                        p.rate = rate_dict[feature[1]]
                product_list.append(p)
    product_dict = {}
    for prod in product_list:
        name = prod.version + ':' + prod.memory + ':' + prod.color + ':' + prod.rate
        if name in product_dict:
            product_dict[name] = product_dict[name] + 1
        else:
            product_dict[name] = 1
    return product_dict
