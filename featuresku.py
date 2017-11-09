import pymysql as db
import xlwt
import numpy as np
import datetime as dt
import configparser


class Product:
    def __init__(self, name, props):
        self.name = name
        self.props = props

cf = configparser.ConfigParser()
cf.read('conf.conf')
dbhost = cf.get('db', 'db_host')
dbuser = cf.get('db', 'db_user')
dbport = cf.getint('db', 'db_port')
dbpass = cf.get('db', 'db_pass')
dbase = cf.get('db', 'db_db')
scon = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, charset='utf8')
scur = scon.cursor()
wb = xlwt.Workbook()

dst_host = cf.get('test', 'host')
dst_user = cf.get('test', 'user')
dst_pass = cf.get('test', 'passwd')
dst_port = cf.getint('test', 'port')
dst_db = cf.get('test', 'db')
dcon = db.connect(host=dst_host, user=dst_user, passwd=dst_pass, db=dst_db, port=dst_port, charset='utf8')
dcur = dcon.cursor()

# warehouse = {1: '分拾', 2: '检测', 3: '市场', 4: '上架', 5: '维修', 6: '报废', 7: 'B端', 8: '预上架', 9: '外包维修',
#              11: '京东', 12: '待卖'}
warehouse = {1: '分拾'}

try:
    versionsql = '''
    SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
    WHERE ppv.pnid = 5
    '''
    vd = {}
    scur.execute(versionsql)
    versions = scur.fetchall()
    for v in versions:
        vd[str(v[0])] = v[1]

    colorsql = '''
    SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
    WHERE ppv.pnid = 10
    '''
    cd = {}
    scur.execute(colorsql)
    colors = scur.fetchall()
    for c in colors:
        cd[str(c[0])] = c[1]

    memorysql = '''
    SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
    WHERE ppv.pnid = 11
    '''
    md = {}
    scur.execute(memorysql)
    memorys = scur.fetchall()
    for m in memorys:
        md[str(m[0])] = m[1]

    prop_sql = '''
    SELECT pm.`model_name`,sw.key_props FROM panda.`stg_warehouse` sw
    LEFT JOIN panda.`pdi_model` pm
    ON sw.`model_id` = pm.`model_id`
    LEFT JOIN panda.`pdi_product` pp
    ON sw.`product_id` = pp.product_id
    WHERE sw.`warehouse_status` = 1
    '''
    scur.execute(prop_sql)
    result = scur.fetchall()
    pset = set()
    for r in result:
        product = Product(r[0], r[1])
        props = product.props.split(';')
        for feature in props:
            f = feature.split(':')
            if f[0] == '5':
                product.version = vd[f[1]]
            if f[0] == '10':
                product.color = cd[f[1]]
            if f[0] == '11':
                product.memory = md[f[1]]
        pname = product.version + ':' + product.color + ':' + product.memory
        pset.add(pname)

    conditionlt1 = 'bt.times < 1 '
    conditionlt3 = 'bt.times > 1 and bt.times <3 '
    conditionlt7 = 'bt.times > 3 and bt.times <7 '
    conditionlt15 = 'bt.times > 7 and bt.times <15 '
    conditionlt30 = 'bt.times > 15 and bt.times <30 '
    conditiongt30 = 'bt.times > 30 '

    for wnum in warehouse:
        plist = [x for x in pset]

        oneday = np.zeros(len(plist), dtype=int)
        threeday = np.zeros(len(plist), dtype=int)
        sevenday = np.zeros(len(plist), dtype=int)
        fifteenday = np.zeros(len(plist), dtype=int)
        thirtyday = np.zeros(len(plist), dtype=int)
        outthirtyday = np.zeros(len(plist), dtype=int)

        for prod in pset:
            v, c, m = prod.split(':')
            cnum = [x[0] for x in cd.items() if c in x[1]]
            mnum = [x[0] for x in md.items() if m in x[1]]
            count_sql = '''
                SELECT bt.model_name,COUNT(1) FROM
                (
                SELECT psw.model_name,(UNIX_TIMESTAMP(NOW())-UNIX_TIMESTAMP(psw.change_time))/60/60/24 `times` FROM
                (
                SELECT sw.*, pm.model_name ,pp.`buy_at` FROM panda.`stg_warehouse` sw
                LEFT JOIN panda.`pdi_model` pm
                ON sw.`model_id` = pm.`model_id`
                LEFT JOIN panda.`pdi_product` pp
                ON sw.`product_id` = pp.`product_id`
                WHERE sw.`warehouse_status` = 1
                AND sw.`warehouse_num` = {}
                AND pm.`model_name` LIKE '%{}%'
                AND pp.`key_props` LIKE '%10:{}%'
                AND pp.`key_props` LIKE '%11:{}%'
                ) psw
                )bt
                WHERE {}
            '''
            number = scur.execute(count_sql.format(wnum, v, cnum[0], mnum[0], conditionlt1))
            if number > 0:
                result = scur.fetchone()
                oneday[plist.index(prod)] = result[1]

            number = scur.execute(count_sql.format(wnum, v, cnum[0], mnum[0], conditionlt3))
            if number > 0:
                result = scur.fetchone()
                threeday[plist.index(prod)] = result[1]

            number = scur.execute(count_sql.format(wnum, v, cnum[0], mnum[0], conditionlt7))
            if number > 0:
                result = scur.fetchone()
                sevenday[plist.index(prod)] = result[1]

            number = scur.execute(count_sql.format(wnum, v, cnum[0], mnum[0], conditionlt15))
            if number > 0:
                result = scur.fetchone()
                fifteenday[plist.index(prod)] = result[1]

            number = scur.execute(count_sql.format(wnum, v, cnum[0], mnum[0], conditionlt30))
            if number > 0:
                result = scur.fetchone()
                thirtyday[plist.index(prod)] = result[1]

            number = scur.execute(count_sql.format(wnum, v, cnum[0], mnum[0], conditiongt30))
            if number > 0:
                result = scur.fetchone()
                outthirtyday[plist.index(prod)] = result[1]

        plist.insert(0, '周期')
        plist.append('总计')
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

        matrix = [plist, l1, l3, l7, l15, l30, g30]
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
        dcur.execute(tablesql.format(wnum, wnum))
        dcon.commit()

        dst_arg = []
        for row in matrix4:
            if row[0] == '周期':
                continue
            if int(row[1]) == 0 and int(row[2]) == 0 and int(row[3]) == 0 and int(row[4]) == 0 and int(row[5]) == 0 \
                    and int(row[6]) == 0:
                continue
            dst_arg.append((str(row[0]), int(row[1]), int(row[2]), int(row[3]), int(row[4]), int(row[5]), int(row[6]),
                            int(row[1]) + int(row[2]) + int(row[3]) + int(row[4]) + int(row[5]) + int(row[6])))

        insert = '''
        insert into ods.ods_product_warehouse_{} VALUES (%s,%s,%s,%s,%s,%s,%s,%s)'''
        dcur.executemany(insert.format(wnum), dst_arg)
        dcon.commit()

        read_sql = '''
        select * from ods.ods_product_warehouse_{}
        '''
        dcur.execute(read_sql.format(wnum))
        result = dcur.fetchall()
        sheet = wb.add_sheet(warehouse[wnum])
        sheet.write(0, 0, '型号')
        sheet.write(0, 1, '颜色')
        sheet.write(0, 2, '内存')
        sheet.write(0, 3, '小于1天内')
        sheet.write(0, 4, '小于3天内')
        sheet.write(0, 5, '小于7天内')
        sheet.write(0, 6, '小于15天内')
        sheet.write(0, 7, '小于30天内')
        sheet.write(0, 8, '大于30天')
        for i, r in enumerate(result):
            if r[0] == '总计':
                sheet.write(i+1, 0, r[0])
            else:
                v, c, m = r[0].split(':')
                sheet.write(i+1, 0, v)
                sheet.write(i+1, 1, c)
                sheet.write(i+1, 2, m)
            sheet.write(i+1, 3, r[1])
            sheet.write(i+1, 4, r[2])
            sheet.write(i+1, 5, r[3])
            sheet.write(i+1, 6, r[4])
            sheet.write(i+1, 7, r[5])
            sheet.write(i+1, 8, r[6])
    path = cf.get('path', 'path')
    wb.save(path + str(dt.datetime.now().date()) + '.xls')
finally:
    dcur.close()
    dcon.close()
    scur.close()
    scon.close()
