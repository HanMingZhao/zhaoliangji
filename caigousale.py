import pymysql as db
import numpy as np
import xlwt
import time
import config as conf

warehousenums = {3: '市场', 12: '待卖'}
cf = conf.product
src_con = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], db=cf['db'],
                     charset=conf.char)
src_cur = src_con.cursor()
cf = conf.test
dst_con = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], db=cf['db'],
                     charset=conf.char)
dst_cur = dst_con.cursor()
start_time = time.time()
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
for whnum in warehousenums:
    oneday = np.zeros(len(sku), dtype=int)
    threeday = np.zeros(len(sku), dtype=int)
    sevenday = np.zeros(len(sku), dtype=int)
    fifteenday = np.zeros(len(sku), dtype=int)
    thirtyday = np.zeros(len(sku), dtype=int)
    outthirtyday = np.zeros(len(sku), dtype=int)

    sqli = '''SELECT bt.model_name,COUNT(1) FROM
    (
    SELECT pm.model_name,psw.product_id,psw.product_name,psw.imei,
    IF(psw.change_time>0,(UNIX_TIMESTAMP(NOW())-UNIX_TIMESTAMP(psw.change_time))/60/60/24,
    (UNIX_TIMESTAMP(NOW())-UNIX_TIMESTAMP(psw.in_time))/60/60/24) `times` FROM 
    (
    SELECT sw.*,pp.`buy_at` FROM panda.`stg_warehouse` sw
    LEFT JOIN panda.`pdi_product` pp ON sw.`product_id` = pp.`product_id`
    ) psw
    LEFT JOIN panda.`pdi_model`  pm  ON  psw.model_id = pm.model_id
    WHERE psw.warehouse_status = 1
    AND psw.warehouse_num = {}
    AND pm.model_name LIKE '%iphone%'
    ORDER BY pm.model_name ,times,imei
    )bt
    WHERE {} 
    GROUP BY bt.model_name'''

    src_cur.execute(sqli.format(whnum, 'bt.times < 1 '))
    results = src_cur.fetchall()
    for r in results:
        oneday[sku.index(r[0])] = r[1]

    src_cur.execute(sqli.format(whnum, 'bt.times > 1 and bt.times <3 '))
    results = src_cur.fetchall()
    for r in results:
        threeday[sku.index(r[0])] = r[1]

    src_cur.execute(sqli.format(whnum, 'bt.times > 3 and  bt.times < 7'))
    results = src_cur.fetchall()
    for r in results:
        sevenday[sku.index(r[0])] = r[1]

    src_cur.execute(sqli.format(whnum, 'bt.times > 7 and bt.times < 15 '))
    results = src_cur.fetchall()
    for r in results:
        fifteenday[sku.index(r[0])] = r[1]

    src_cur.execute(sqli.format(whnum, 'bt.times > 15 and bt.times < 30 '))
    results = src_cur.fetchall()
    for r in results:
        thirtyday[sku.index(r[0])] = r[1]

    src_cur.execute(sqli.format(whnum, 'bt.times > 30 '))
    results = src_cur.fetchall()
    for r in results:
        outthirtyday[sku.index(r[0])] = r[1]

    sqla = '''SELECT bt.brand_name,COUNT(1) FROM 
    (
    SELECT pm.brand_name,psw.product_id,psw.product_name,psw.imei,
    IF(psw.change_time>0,(UNIX_TIMESTAMP(NOW())-UNIX_TIMESTAMP(psw.change_time))/60/60/24,
    (UNIX_TIMESTAMP(NOW())-UNIX_TIMESTAMP(psw.in_time))/60/60/24) `times` FROM
    (
    SELECT sw.*,pp.`buy_at` FROM panda.`stg_warehouse` sw
    LEFT JOIN panda.`pdi_product` pp ON sw.`product_id` = pp.`product_id`
    ) psw
    LEFT JOIN panda.`pdi_model`  pm  ON  psw.model_id = pm.model_id
    WHERE psw.warehouse_status = 1
    AND psw.warehouse_num = {}
    AND pm.model_name NOT LIKE '%iphone%'
    ORDER BY pm.brand_name ,times,imei
    )bt
    WHERE {}
    GROUP BY bt.brand_name'''

    src_cur.execute(sqla.format(whnum, 'bt.times < 1 '))
    results = src_cur.fetchall()
    for r in results:
        oneday[sku.index(r[0])] = r[1]

    src_cur.execute(sqla.format(whnum, 'bt.times > 1 and bt.times < 3 '))
    results = src_cur.fetchall()
    for r in results:
        threeday[sku.index(r[0])] = r[1]

    src_cur.execute(sqla.format(whnum, 'bt.times > 3 and bt.times < 7 '))
    results = src_cur.fetchall()
    for r in results:
        sevenday[sku.index(r[0])] = r[1]

    src_cur.execute(sqla.format(whnum, 'bt.times > 7 and bt.times < 15 '))
    results = src_cur.fetchall()
    for r in results:
        fifteenday[sku.index(r[0])] = r[1]

    src_cur.execute(sqla.format(whnum, 'bt.times > 15 and bt.times < 30 '))
    results = src_cur.fetchall()
    for r in results:
        thirtyday[sku.index(r[0])] = r[1]

    src_cur.execute(sqla.format(whnum, 'bt.times > 30 '))
    results = src_cur.fetchall()
    for r in results:
        outthirtyday[sku.index(r[0])] = r[1]

    sku.insert(0, '周期')
    sku.append('总计')
    l1 = [int(x) for x in oneday]
    l1.append(sum(l1))
    l3 = [int(x) for x in threeday]
    l3.append(sum(l3))
    l7 = [int(x) for x in sevenday]
    l7.append(sum(l7))
    l15 = [int(x) for x in fifteenday]
    l15.append(sum(l15))
    l30 = [int(x) for x in thirtyday]
    l30.append(sum(l30))
    g30 = [int(x) for x in outthirtyday]
    g30.append(sum(g30))
    l1.insert(0, '小于1天内')
    l3.insert(0, '小于3天内')
    l7.insert(0, '小于7天内')
    l15.insert(0, '小于15天内')
    l30.insert(0, '小于30天内')
    g30.insert(0, '大于30天')

    matrix = [sku, l1, l3, l7, l15, l30, g30]
    matrix2 = np.matrix(matrix)
    matrix3 = matrix2.transpose()
    matrix4 = matrix3.tolist()

    tablesql = '''drop table if EXISTS ods.ods_product_ez_warehouse_{} ;
    create TABLE if NOT EXISTS ods.ods_product_ez_warehouse_{} 
    ( `sku` VARCHAR(32), 
    `小于1天内` SMALLINT(4), 
    `小于3天内` SMALLINT(4), 
    `小于7天内` SMALLINT(4), 
    `小于15天内` SMALLINT(4), 
    `小于30天内` SMALLINT(4), 
    `大于30天` SMALLINT(4),
    `总计` SMALLINT(4)
    ) 
    ENGINE=MYISAM CHARSET=utf8; 
    '''
    dst_cur.execute(tablesql.format(whnum, whnum))
    dst_con.commit()

    dst_arg = []
    for row in matrix4:
        if row[0] == '周期':
            continue
        if int(row[1]) == 0 and int(row[2]) == 0 and int(row[3]) == 0 and int(row[4]) == 0 \
                and int(row[5]) == 0 and int(row[6]) == 0:
            continue
        dst_arg.append((str(row[0]), int(row[1]), int(row[2]), int(row[3]), int(row[4]), int(row[5]), int(row[6]),
                        int(row[1]) + int(row[2]) + int(row[3]) + int(row[4]) + int(row[5]) + int(row[6])))

    insert = '''
    insert into ods.ods_product_ez_warehouse_{} VALUES (%s,%s,%s,%s,%s,%s,%s,%s)'''
    dst_cur.executemany(insert.format(whnum), dst_arg)
    dst_con.commit()

    sheet = wb.add_sheet(warehousenums[whnum])
    sheet.write(0, 0, 'sku')
    sheet.write(0, 1, '小于1天内')
    sheet.write(0, 2, '小于3天内')
    sheet.write(0, 3, '小于7天内')
    sheet.write(0, 4, '小于15天内')
    sheet.write(0, 5, '小于30天内')
    sheet.write(0, 6, '大于30天')
    sheet.write(0, 7, '总计')

    read_sql = '''
    select * from ods.ods_product_ez_warehouse_{}
    '''
    dst_cur.execute(read_sql.format(whnum))
    result = dst_cur.fetchall()
    for i, r in enumerate(result):
        sheet.write(i+1, 0, r[0])
        sheet.write(i+1, 1, r[1])
        sheet.write(i+1, 2, r[2])
        sheet.write(i+1, 3, r[3])
        sheet.write(i+1, 4, r[4])
        sheet.write(i+1, 5, r[5])
        sheet.write(i+1, 6, r[6])
        sheet.write(i+1, 7, r[7])
    print('runtime：', time.time()-start_time)

wb.save(conf.path + conf.today.strftime(conf.date_format) + 'market.xls')

src_cur.close()
src_cur.close()
