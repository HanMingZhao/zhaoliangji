import pymysql as db
import config as conf
import xlwt


def supplier_sum(db_cursor, workbook, sheet_name, start, end):
    in_sql = '''
    SELECT ps.`name`,COUNT(1) FROM panda.`pdi_product_cost` ppc
    LEFT JOIN PANDA.`pdi_suppiler` ps
    ON ppc.`supplier` = ps.`suppiler_id`
    WHERE ppc.`created_at` >'{}'
    AND ppc.`created_at` < '{}'
    GROUP BY ppc.`supplier`
    '''
    db_cursor.execute(in_sql.format(start, end))
    result = cursor.fetchall()
    sheet = workbook.add_sheet(sheet_name)
    sheet.write(0, 0, '供应商')
    sheet.write(0, 1, '供货数量')
    for i, r in enumerate(result):
        sheet.write(i+1, 0, r[0])
        sheet.write(i+1, 1, r[1])

cf = conf.product
connect = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cursor = connect.cursor()
wb = xlwt.Workbook()

supplier_sum(cursor, wb, '全年', '2017-1-1', '2018-1-1')
supplier_sum(cursor, wb, '1-5月', '2017-1-1', '2017-6-1')
supplier_sum(cursor, wb, '6-10月', '2017-6-1', '2017-11-1')
supplier_sum(cursor, wb, '11-12月', '2017-11-1', '2018-1-1')

wb.save(conf.path + '2017supplier.xls')
cursor.close()
connect.close()
