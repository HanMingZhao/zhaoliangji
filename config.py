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

target = 5500

path = '/var/www/python/'

charset = 'utf8'

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
                    if feature[0] == '5' and version_dict is not None:
                        p.version = version_dict[feature[1]]
                    if feature[0] == '10' and color_dict is not None:
                        p.color = color_dict[feature[1]]
                    if feature[0] == '11' and memory_dict is not None:
                        p.memory = memory_dict[feature[1]]
                    if feature[0] == '12' and rate_dict is not None:
                        p.rate = rate_dict[feature[1]]
                if hasattr(p, 'version') and hasattr(p, 'memory') and hasattr(p, 'color') and hasattr(p, 'rate'):
                    product_list.append(p)
    product_dict = collections.OrderedDict()
    for prod in product_list:
        name = prod.version + ':' + prod.memory if hasattr(prod, 'memory') else '' + ':' + prod.color \
            if hasattr(prod, 'color') else '' + ':' + prod.rate if hasattr(prod, 'rate') else ''
        if name in product_dict:
            product_dict[name] = product_dict[name] + 1
        else:
            product_dict[name] = 1
    return product_dict

version_memory_dict = {4: {21, 23, 24}, 5: {21, 22, 23}, 14: {21, 22, 23, 24}, 15: {21, 22, 23, 24}, 16: {21, 23, 24},
                       114: {21, 22, 23, 24}, 127: {22, 24, 128}, 129: {22, 24, 128}, 525: {23, 128}, 544: {23, 128},
                       618: {23, 128}}
