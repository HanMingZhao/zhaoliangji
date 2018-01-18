import config as conf
import xlwt


def sale_product(workbook, mbrand, mtype, sheet_name, version_dict, memory_dict, color_dict):
    sku_sql = '''
    SELECT sws.key_props FROM panda.`stg_warning_sku` sws
    WHERE sws.type_id = {}
    AND sws.brand_id= {}
    '''
    conf.product_cursor.execute(sku_sql.format(mtype, mbrand))
    sku_result = conf.product_cursor.fetchall()
    sku_dict = conf.product_count(sku_result, version_dict, memory_dict, color_dict)

    sale_sql = '''
    SELECT pp.`key_props`,pp.`product_name` FROM panda.`odi_order` oo
    LEFT JOIN panda.`pdi_product` pp
    ON oo.`product_id` = pp.`product_id`
    WHERE oo.`order_status` IN (1,2,4,5)
    AND oo.`order_type` IN (1,2)
    AND pp.`brand_id` = {}
    AND pp.`type_id` = {}
    AND oo.`pay_at` > '{}'
    AND oo.`pay_at` < '{}'
    AND pp.`product_name` NOT LIKE '测试%'
    AND pp.`product_name` NOT LIKE '配件%'
    '''
    conf.product_cursor.execute(sale_sql.format(mbrand, mtype, conf.last_week_day.strftime(conf.date_format),
                                                conf.today.strftime(conf.date_format)))
    sale_result = conf.product_cursor.fetchall()
    sale_dict = conf.product_count(sale_result, version_dict, memory_dict, color_dict)

    store_sql ='''
    SELECT sw.key_props FROM panda.`stg_warehouse` sw
    WHERE sw.warehouse_num IN (1,2,4,7)
    AND sw.warehouse_status = 1
    AND sw.brand_id = {}
    '''
    conf.product_cursor.execute(store_sql.format(mbrand))
    store_result = conf.product_cursor.fetchall()
    store_dict = conf.product_count(store_result, version_dict, memory_dict, color_dict)

    sheet = workbook.add_sheet(sheet_name)
    sheet.write(0, 0, '型号')
    sheet.write(0, 1, '内存')
    sheet.write(0, 2, '颜色')
    sheet.write(0, 3, '7天销售')
    sheet.write(0, 4, '库存')
    for s in sku_dict:
        pv, pm, pc = s.split(':')
        row = len(sheet.rows)
        sheet.write(row, 0, pv)
        sheet.write(row, 1, pm)
        sheet.write(row, 2, pc)
        sheet.write(row, 3, sale_dict[s] if s in sale_dict else 0)
        sheet.write(row, 4, store_dict[s] if s in store_dict else 0)

    for s in sale_dict:
        if s not in sku_dict:
            pv, pm, pc = s.split(':')
            row = len(sheet.rows)
            sheet.write(row, 0, pv)
            sheet.write(row, 1, pm)
            sheet.write(row, 2, pc)
            sheet.write(row, 3, sale_dict[s])
            sheet.write(row, 4, store_dict[s] if s in store_dict else 0)

    for s in store_dict:
        if s not in sku_dict and s not in sale_dict:
            pv, pm, pc = s.split(':')
            row = len(sheet.rows)
            sheet.write(row, 0, pv)
            sheet.write(row, 1, pm)
            sheet.write(row, 2, pc)
            sheet.write(row, 3, sale_dict[s] if s in sale_dict else 0)
            sheet.write(row, 4, store_dict[s])

wb = xlwt.Workbook()
propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''
vd = conf.properties_dict(conf.product_cursor, propsql, 5)
md = conf.properties_dict(conf.product_cursor, propsql, 11)
cd = conf.properties_dict(conf.product_cursor, propsql, 10)

sale_product(wb, 1, 1, 'iphone', vd, md, cd)
sale_product(wb, 2, 1, 'oppo', vd, md, cd)
sale_product(wb, 3, 1, 'vivo', vd, md, cd)
sale_product(wb, 4, 1, '三星', vd, md, cd)
sale_product(wb, 5, 1, '小米', vd, md, cd)
sale_product(wb, 6, 1, '魅族', vd, md, cd)
sale_product(wb, 7, 1, '华为', vd, md, cd)
sale_product(wb, 8, 1, '一加', vd, md, cd)
sale_product(wb, 10, 1, '锤子', vd, md, cd)
sale_product(wb, 13, 1, '美图', vd, md, cd)
sale_product(wb, 1, 2, 'ipad', vd, md, cd)
sale_product(wb, 1, 3, 'mac', vd, md, cd)
sale_product(wb, 1, 9, '手表', vd, md, cd)
sale_product(wb, 15, 3, '联想', vd, md, cd)
sale_product(wb, 19, 3, '戴尔', vd, md, cd)
sale_product(wb, 20, 3, '华硕', vd, md, cd)


wb.save('sku400.xls')
conf.product_cursor.close()
conf.product_connect.close()
