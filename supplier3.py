import pymysql as db
import config as conf
import xlwt


def query_add_supplier(db_cursor, workbook, start, end, sheet_name):
    grow_sql = '''
    SELECT ppc.supplier,ps.name FROM panda.`pdi_product_cost` ppc
    LEFT JOIN PANDA.`pdi_suppiler` ps
    ON ppc.`supplier` = ps.`suppiler_id`
    WHERE ppc.`created_at` >'2017-{}-1'
    AND ppc.`created_at` < '201{}-1'
    AND ppc.supplier NOT IN 
    (
    SELECT ppc.supplier FROM panda.`pdi_product_cost` ppc
    LEFT JOIN PANDA.`pdi_suppiler` ps
    ON ppc.`supplier` = ps.`suppiler_id`
    WHERE ppc.`created_at` >'2017-1-1'
    AND ppc.`created_at` < '2017-{}-1'
    GROUP BY ppc.`supplier`
    )
    GROUP BY ppc.`supplier`
    '''
    db_cursor.execute(grow_sql.format(start, end, start))
    result = db_cursor.fetchall()
    sheets = workbook.add_sheet(sheet_name)
    sheets.write(0, 0, '供应商ID')
    sheets.write(0, 1, '供应商')
    for j, r in enumerate(result):
        sheets.write(j + 1, 0, r[0])
        sheets.write(j + 1, 1, r[1])
    return None

cf = conf.product
connect = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cursor = connect.cursor()
wb = xlwt.Workbook()
init_sql = '''
SELECT ppc.supplier,ps.name FROM panda.`pdi_product_cost` ppc
LEFT JOIN PANDA.`pdi_suppiler` ps
ON ppc.`supplier` = ps.`suppiler_id`
WHERE ppc.`created_at` >'2017-1-1'
AND ppc.`created_at` < '2017-6-1'
GROUP BY ppc.`supplier`
'''
cursor.execute(init_sql)
init_result = cursor.fetchall()
sheet = wb.add_sheet('1-5')
sheet.write(0, 0, '供应商ID')
sheet.write(0, 1, '供应商')
for i, res in enumerate(init_result):
    sheet.write(i+1, 0, res[0])
    sheet.write(i+1, 1, res[1])

query_add_supplier(cursor, wb, '6', '7-11', '6-10')
query_add_supplier(cursor, wb, '11', '8-1', '11-12')

wb.save('supplier3.xls')
cursor.close()
connect.close()
