import pymysql as db
import numpy as np
import xlwt
import datetime as dt

warehouse_nums = [1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12]

src_con = db.connect(host='rm-bp13wnvyc2dh86ju1.mysql.rds.aliyuncs.com', user='panda_reader', passwd='zhaoliangji3503',
                     db='panda', charset='utf8')
dst_con = db.connect(host='114.215.176.190', user='root', passwd='huodao123', db='ods', port=33069, charset='utf8')

src_cur = src_con.cursor()
dst_cur = dst_con.cursor()

wb = xlwt.Workbook()

sku = []

modelsql = '''SELECT DISTINCT(pm.model_name) FROM
(
SELECT sw.*,pp.`buy_at` FROM panda.`stg_warehouse` sw
LEFT JOIN panda.`pdi_product` pp ON sw.`product_id` = pp.`product_id`
) psw
LEFT JOIN panda.`pdi_model`  pm  ON  psw.model_id = pm.model_id
WHERE psw.warehouse_status = 1
AND pm.model_name LIKE '%iphone%'
ORDER BY pm.model_name '''
src_cur.execute(modelsql)
models = src_cur.fetchall()
for m in models:
    sku.append(m[0])

brandsql = '''SELECT DISTINCT(pm.brand_name) FROM
(
SELECT sw.*,pp.`buy_at` FROM panda.`stg_warehouse` sw
LEFT JOIN panda.`pdi_product` pp ON sw.`product_id` = pp.`product_id`
) psw
LEFT JOIN panda.`pdi_model`  pm  ON  psw.model_id = pm.model_id
WHERE psw.warehouse_status = 1
AND pm.model_name NOT LIKE '%iphone%'
ORDER BY pm.brand_name'''
src_cur.execute(brandsql)
brands = src_cur.fetchall()
for b in brands:
    sku.append(b[0])

w1 = np.zeros(len(sku), dtype=int)
w2 = np.zeros(len(sku), dtype=int)
w3 = np.zeros(len(sku), dtype=int)
w4 = np.zeros(len(sku), dtype=int)
w5 = np.zeros(len(sku), dtype=int)
w6 = np.zeros(len(sku), dtype=int)
w7 = np.zeros(len(sku), dtype=int)
w8 = np.zeros(len(sku), dtype=int)
w9 = np.zeros(len(sku), dtype=int)
w11 = np.zeros(len(sku), dtype=int)
wdict = {1: w1, 2: w2, 3: w3, 4: w4, 5: w5, 6: w6, 7: w7, 8: w8, 9: w9, 11: w11}

for whnum in warehouse_nums:
    askusql = '''SELECT pm.brand_name,COUNT(1) FROM panda.`stg_warehouse` sw 
    LEFT JOIN panda.`pdi_model` pm ON sw.model_id = pm.model_id
    WHERE sw.warehouse_status = 1
    AND sw.warehouse_num = {} 
    AND pm.model_name NOT LIKE '%iphone%'
    GROUP BY pm.brand_name,sw.warehouse_num'''
    src_cur.execute(askusql.format(whnum))
    results = src_cur.fetchall()
    for r in results:
        wdict[whnum][sku.index(r[0])] = r[1]

    iskusql = '''SELECT pm.model_name,COUNT(1) FROM panda.`stg_warehouse` sw 
    LEFT JOIN panda.`pdi_model` pm ON sw.model_id = pm.model_id
    WHERE sw.warehouse_status = 1
    AND sw.warehouse_num ={} 
    AND pm.model_name LIKE '%iphone%'
    GROUP BY pm.model_name,sw.warehouse_num'''
    src_cur.execute(iskusql.format(whnum))
    results = src_cur.fetchall()
    for r in results:
        wdict[whnum][sku.index(r[0])] = r[1]

sku.insert(0, '库位')
sku.append('总计')
l1 = [int(x) for x in w1]
l1.append(sum(l1))
l2 = [int(x) for x in w2]
l2.append(sum(l2))
l3 = [int(x) for x in w3]
l3.append(sum(l3))
l4 = [int(x) for x in w4]
l4.append(sum(l4))
l5 = [int(x) for x in w5]
l5.append(sum(l5))
l6 = [int(x) for x in w6]
l6.append(sum(l6))
l7 = [int(x) for x in w7]
l7.append(sum(l7))
l8 = [int(x) for x in w8]
l8.append(sum(l8))
l9 = [int(x) for x in w9]
l9.append(sum(l9))
l11 = [int(x) for x in w11]
l11.append(sum(l11))
l1.insert(0, '分拾')
l2.insert(0, '检测')
l3.insert(0, '市场')
l4.insert(0, '上架')
l5.insert(0, '维修')
l6.insert(0, '报废')
l7.insert(0, 'B端')
l8.insert(0, '预上架')
l9.insert(0, '外包维修')
l11.insert(0, '京东')

matrix = np.matrix([sku, l1, l2, l3, l5, l6, l7, l8, l9, l11, l4])
matrix2 = matrix.transpose().tolist()

tablesql = '''
drop table if EXISTS ods.ods_warehouse_sum;
CREATE TABLE ods.ods_warehouse_sum 
(`sku` VARCHAR(32), 
`分拾` SMALLINT(4), 
`检测` SMALLINT(4), 
`市场` SMALLINT(4), 
`维修` SMALLINT(4), 
`报废` SMALLINT(4), 
`B端` SMALLINT(4), 
`预上架` SMALLINT(4), 
`外包维修` SMALLINT(4), 
`京东` SMALLINT(4),
`上架` SMALLINT(4),
`总计` SMALLINT(4)
)ENGINE=MYISAM CHARSET=utf8; 
'''
dst_cur.execute(tablesql)
dst_con.commit()

dst_arg = []
for row in matrix2:
    if row[0] == '库位':
        continue
    if row[0] == '总计' and int(row[1]) == 0 and int(row[2]) == 0 and int(row[3]) == 0 and int(row[4]) == 0\
            and int(row[5]) == 0 and int(row[6]) == 0 and int(row[7]) == 0 and int(row[8]) == 0 and int(row[9]) == 0\
            and int(row[10]) == 0:
        continue
    dst_arg.append((row[0], int(row[1]), int(row[2]), int(row[3]), int(row[4]), int(row[5]), int(row[6]), int(row[7]),
                    int(row[8]), int(row[9]), int(row[10]), int(row[1]) + int(row[2]) + int(row[3]) + int(row[4]) +
                    int(row[5]) + int(row[6]) + int(row[7]) + int(row[8]) + int(row[9]) + int(row[10])))

insertsql = '''insert into ods.ods_warehouse_sum  VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s)'''

dst_cur.executemany(insertsql, dst_arg)
dst_con.commit()

sheet = wb.add_sheet('sheet1')
sheet.write(0, 0, 'sku')
sheet.write(0, 1, '分拾')
sheet.write(0, 2, '检测')
sheet.write(0, 3, '市场')
sheet.write(0, 4, '维修')
sheet.write(0, 5, '报废')
sheet.write(0, 6, 'B端')
sheet.write(0, 7, '预上架')
sheet.write(0, 8, '外包维修')
sheet.write(0, 9, '京东')
sheet.write(0, 10, '上架')
sheet.write(0, 11, '总计')
sheet.write(0, 12, '占比')
querySql = '''
select * from ods.ods_warehouse_sum
'''
dst_cur.execute(querySql)
result = dst_cur.fetchall()
for i, r in enumerate(result):
    for j, x in enumerate(r):
        sheet.write(i+1, j, int(x) if j > 0 else x)
    sheet.write(i+1, len(r), int(r[10])/int(r[11]))
sheetLength = len(sheet.rows)
lastRow = len(result) - 1
sheet.write(sheetLength, 0, '占比')
sheet.write(sheetLength, 1, int(result[lastRow][1])/int(result[lastRow][11]))
sheet.write(sheetLength, 2, int(result[lastRow][2])/int(result[lastRow][11]))
sheet.write(sheetLength, 3, int(result[lastRow][3])/int(result[lastRow][11]))
sheet.write(sheetLength, 4, int(result[lastRow][4])/int(result[lastRow][11]))
sheet.write(sheetLength, 5, int(result[lastRow][5])/int(result[lastRow][11]))
sheet.write(sheetLength, 6, int(result[lastRow][6])/int(result[lastRow][11]))
sheet.write(sheetLength, 7, int(result[lastRow][7])/int(result[lastRow][11]))
sheet.write(sheetLength, 8, int(result[lastRow][8])/int(result[lastRow][11]))
sheet.write(sheetLength, 9, int(result[lastRow][9])/int(result[lastRow][11]))
sheet.write(sheetLength, 10, int(result[lastRow][10])/int(result[lastRow][11]))

path = '/var/www/python3/'
wb.save(path + str(dt.datetime.today().date()) + 'warehouse.xls')

src_cur.close()
src_con.close()
dst_cur.close()
dst_con.close()

print('done!\n'*5)

