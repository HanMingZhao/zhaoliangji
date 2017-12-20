import pymysql
import config
import xlwt

cf = config.product
src_con = pymysql.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], db=cf['db'],
                          charset=config.charset)
src_cur = src_con.cursor()
workbook = xlwt.Workbook()

in_sql = '''
SELECT ps.`name`,COUNT(1) FROM panda.`pdi_product_cost` ppc
LEFT JOIN PANDA.`pdi_suppiler` ps
ON ppc.`supplier` = ps.`suppiler_id`
WHERE ppc.`created_at` >'2017-11-1'
AND ppc.`created_at` < '2017-12-1'
GROUP BY ppc.`supplier`
'''
src_cur.execute(in_sql)
result = src_cur.fetchall()
suppliers = set()
in_dict = {}
for r in result:
    in_dict[r[0]] = r[1]
    suppliers.add(r[0])

out_sql = '''
SELECT ps.name,COUNT(1) FROM panda.`pdi_suppiler_track` pst
LEFT JOIN panda.`pdi_suppiler` ps
ON pst.`suppiler_id` = ps.`suppiler_id`
WHERE pst.`track_type`=2
AND pst.`created_at` >'2017-11-1'
AND pst.`created_at` < '2017-12-1'
GROUP BY pst.`suppiler_id`
'''
src_cur.execute(out_sql)
result = src_cur.fetchall()
out_dict = {}
for r in result:
    out_dict[r[0]] = r[1]
    suppliers.add(r[0])

sheet = workbook.add_sheet('sheet')
sheet.write(0, 0, '供应商')
sheet.write(0, 1, '进货')
sheet.write(0, 0, '退货')
for i, s in enumerate(suppliers):
    sheet.write(i+1, 0, s)
    sheet.write(i+1, 1, in_dict[s] if s in in_dict else 0)
    sheet.write(i+1, 2, out_dict[s] if s in out_dict else 0)

workbook.save(config.path + 'supplier.xls')
src_cur.close()
src_con.close()
