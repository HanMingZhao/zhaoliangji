import pymysql as db
import config as conf
import xlwt


def product_count(sql_results, version_dict, memory_dict, color_dict):
    product_dict = {}
    for sr in sql_results:
        if sr[0] is not None:
            p = conf.Product(sr[0])
            properties = p.props.split(';')
            for f in properties:
                feature = f.split(':')
                if feature[0] == '5':
                    p.version = version_dict[feature[1]]
                if feature[0] == '10':
                    p.color = color_dict[feature[1]]
                if feature[0] == '11':
                    p.memory = memory_dict[feature[1]]
            name = p.version + ':' + p.memory + ':' + p.color
            product_dict[name] = sr[1]
    return product_dict


def sales_sku(cursor, workbook, start, end, sheet):
    sale_sql = '''
    select pp.key_props from panda.odi_order oo
    left join panda.pdi_product pp 
    on oo.product_id = pp.product_id
    where oo.order_status in (1,2,4,5)
    and oo.order_type in (1,2)
    and oo.pay_at < '{}'
    and oo.pay_at > '{}'
    '''
    count = cursor.execute(sale_sql.format(end, start))
    sales_result = cursor.fetchall()
    sales_dict = conf.product_count(sales_result, vd, md, cd)

    sku_sql = '''
    select sws.key_props,sws.sku_id from panda.stg_warning_sku sws
    '''
    cursor.execute(sku_sql)
    sku_result = cursor.fetchall()
    sku_dict = product_count(sku_result, vd, md, cd)

    sheet = workbook.add_sheet(sheet)
    sheet.write(0, 0, 'SKUID')
    sheet.write(0, 1, '型号')
    sheet.write(0, 2, '内存')
    sheet.write(0, 3, '颜色')
    sheet.write(0, 4, '数量')
    sheet.write(0, 5, '占比')
    for i, s in enumerate(sku_dict):
        pv, pm, pc = s.split(':')
        sheet.write(i + 1, 0, sku_dict[s])
        sheet.write(i + 1, 1, pv)
        sheet.write(i + 1, 2, pm)
        sheet.write(i + 1, 3, pc)
        sheet.write(i + 1, 4, sales_dict[s] if s in sales_dict else 0)
        sheet.write(i + 1, 5, sales_dict[s] / count if s in sales_dict else 0)

cf = conf.new_test
connect = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cur = connect.cursor()
wb = xlwt.Workbook()

propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = conf.properties_dict(cur, propsql, 5)
md = conf.properties_dict(cur, propsql, 11)
cd = conf.properties_dict(cur, propsql, 10)

sales_sku(cur, wb, '2017-1-1', '2017-2-1', '1')
sales_sku(cur, wb, '2017-2-1', '2017-3-1', '2')
sales_sku(cur, wb, '2017-3-1', '2017-4-1', '3')
sales_sku(cur, wb, '2017-4-1', '2017-5-1', '4')
sales_sku(cur, wb, '2017-5-1', '2017-6-1', '5')
sales_sku(cur, wb, '2017-6-1', '2017-7-1', '6')
sales_sku(cur, wb, '2017-7-1', '2017-8-1', '7')
sales_sku(cur, wb, '2017-8-1', '2017-9-1', '8')
sales_sku(cur, wb, '2017-9-1', '2017-10-1', '9')
sales_sku(cur, wb, '2017-10-1', '2017-11-1', '10')
sales_sku(cur, wb, '2017-11-1', '2017-12-1', '11')
sales_sku(cur, wb, '2017-12-1', '2018-1-1', '12')


wb.save('2017salesku.xls')
cur.close()
connect.close()
