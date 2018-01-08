import pymysql as db
import config as conf
import datetime
import xlwt
import requests as re


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


def sales_sku(cursor, workbook, start, end, sale_condition, sku_condition, sheet, high_level, mid_level):
    count_sql = '''
    select count(1) from panda.odi_order oo 
    where oo.order_status in (1,2,4,5) 
    and oo.order_type in (1,2) 
    and oo.pay_at > '{}' 
    and oo.pay_at < '{}' 
    '''
    cursor.execute(count_sql.format(start, end))
    count_result = cursor.fetchone()
    count = count_result[0]

    sale_sql = '''
    select pp.key_props from panda.odi_order oo
    left join panda.pdi_product pp 
    on oo.product_id = pp.product_id
    left join panda.pdi_model pm
    on pp.model_id = pm.model_id
    where oo.order_status in (1,2,4,5)
    and oo.order_type in (1,2)
    and oo.pay_at < '{}'
    and oo.pay_at > '{}'
    and {}
    '''
    cursor.execute(sale_sql.format(end, start, sale_condition))
    sales_result = cursor.fetchall()
    sales_dict = conf.product_count(sales_result, vd, md, cd)

    sku_sql = '''
    select sws.key_props,sws.sku_id from panda.stg_warning_sku sws
    where {}
    '''
    cursor.execute(sku_sql.format(sku_condition))
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
        if s in sales_dict:
            level = sales_dict[s]/count
            if level > high_level:
                resp = re.get(conf.warning_sku.format(sku_dict[s], 1))
                print(s, level, conf.warning_sku.format(sku_dict[s], 1), resp)
            elif mid_level < level < high_level:
                resp = re.get(conf.warning_sku.format(sku_dict[s], 2))
                print(conf.warning_sku.format(sku_dict[s], 2), resp)
            elif level < mid_level:
                resp = re.get(conf.warning_sku.format(sku_dict[s], 3))
                print(conf.warning_sku.format(sku_dict[s], 2), resp)

cf = conf.product
connect = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cur = connect.cursor()
wb = xlwt.Workbook()

ios_high = 0.02
ios_mid = 0.01

ipad_high = 0.002
ipad_mid = 0.0005

android_high = 0.003
android_mid = 0.001

propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = conf.properties_dict(cur, propsql, 5)
md = conf.properties_dict(cur, propsql, 11)
cd = conf.properties_dict(cur, propsql, 10)

rolling_date = conf.today - datetime.timedelta(15)
start_str = rolling_date.strftime(conf.date_format)
end_str = conf.today.strftime(conf.date_format)
sales_sku(cur, wb, start_str, end_str, 'pm.model_name like \'%iphone%\'', 'sws.sku_name like \'%iphone%\'', 'iphone',
          ios_high, ios_mid)
sales_sku(cur, wb, start_str, end_str, 'pm.pcid=2', 'sws.sku_name like \'%ipad%\'', 'ipad', ipad_high, ipad_mid)
sales_sku(cur, wb, start_str, end_str, 'pm.pcid=1 and pm.model_name not like \'%iphone%\'',
          'sws.sku_name not like \'%iphone%\' and sws.sku_name not like \'%ipad%\'', 'android', android_high,
          android_mid)

wb.save(conf.path + 'last15days.xls')
cur.close()
connect.close()
