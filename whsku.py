import pymysql as db
import numpy as np

'''
if len(sys.argv) > 1:
    try:
        whnum = int(sys.argv[1])
    except :
        print('arg must be number!')
        sys.exit(0)
else:
    print('input warehouse num!')
    sys.exit(0)

warehousenums = [1, 2, 3, 4, 5, 6, 7, 8, 9, 11]
if whnum not in warehousenums:
    print('wrong warehouse num')
    sys.exit(0)
'''
warehousenums = [1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12]

src_con = db.connect(host='rm-bp13wnvyc2dh86ju1.mysql.rds.aliyuncs.com', user='panda_reader', passwd='zhaoliangji3503',
                   db='panda', charset='utf8')
dst_con = db.connect(host='114.215.176.190', user='root', passwd='huodao123', db='ods', port=33069, charset='utf8')

src_cur = src_con.cursor()
dst_cur = dst_con.cursor()
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

    tablesql = '''drop table if EXISTS ods.ods_product_warehouse_{} ;
    create TABLE if NOT EXISTS ods.ods_product_warehouse_{} 
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
        if row[0] == '总计' and int(row[1]) == 0 and int(row[2]) == 0 and int(row[3]) == 0 and int(row[4]) == 0 \
                and int(row[5]) == 0 and int(row[6]) == 0:
            continue
        dst_arg.append((str(row[0]), int(row[1]), int(row[2]), int(row[3]), int(row[4]), int(row[5]), int(row[6]),
                        int(row[1]) + int(row[2]) + int(row[3]) + int(row[4]) + int(row[5]) + int(row[6])))

    insert = '''
    insert into ods.ods_product_warehouse_{} VALUES (%s,%s,%s,%s,%s,%s,%s,%s)'''
    dst_cur.executemany(insert.format(whnum), dst_arg)
    dst_con.commit()

src_cur.close()
src_con.close()
dst_cur.close()
dst_con.close()

print('done！\n'*5)
