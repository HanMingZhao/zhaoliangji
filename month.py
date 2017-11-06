import pymysql as db
import configparser

cf = configparser.ConfigParser()
cf.read('conf.conf')
dbhost = cf.get('db', 'db_host')
dbuser = cf.get('db', 'db_user')
dbport = cf.getint('db', 'db_port')
dbpass = cf.get('db', 'db_pass')
dbase = cf.get('db', 'db_db')
src_con = db.connect(host=dbhost, user=dbuser, passwd=dbpass, db=dbase, charset='utf8')
dst_con = db.connect(host='114.215.176.190', user='root', passwd='huodao123', db='ods', port=33069, charset='utf8')

src_cur = src_con.cursor()
dst_cur = dst_con.cursor()

from_dict = {'ios': ''' AND aui.`from_shop` IN ('','ios') ''', 'h5': ''' and aui.`from_shop` = 'h5' ''',
             'B': ''' AND aui.`from_shop` IN ('Patica','猎趣','趣分期','中捷代购','钱到到','小卖家','趣先享',
             '京东店铺','机密') ''', 'android': ''' AND aui.`from_shop` NOT IN ('Patica','猎趣','趣分期','中捷代购',
             '钱到到','小卖家','趣先享','京东店铺','机密','h5','','ios') '''}

tablesql = '''
           drop table if EXISTS ods.ods_month_info;
           CREATE TABLE ods.`ods_month_info` (
          `sale_date` date DEFAULT NULL,
          `order` int(11) DEFAULT NULL,
          `pay` int(11) DEFAULT NULL,
          `cancel` int(11) DEFAULT NULL,
          `total` decimal(9,2) DEFAULT NULL,
          `from_shop` varchar(32) DEFAULT NULL
          ) ENGINE=MyISAM DEFAULT CHARSET=utf8
           '''
dst_cur.execute(tablesql)
dst_con.commit()

for d in from_dict:
    numsql = '''
    SELECT DATE(oo.create_at) `date`,COUNT(1) `nums`,SUM(oo.total_amount) `total` FROM panda.`odi_order` oo
    LEFT JOIN panda.`aci_user_info` aui
    ON oo.`user_id` = aui.`user_id`
    WHERE oo.`order_status` IN (1,2,4,5)
    AND oo.`create_at` > '2017-8-1 00:00:00'
    {}
    GROUP BY DATE(oo.`create_at`)
    '''

    sumsql = '''
    SELECT DATE(oo.`create_at`) `date`,COUNT(1) FROM panda.`odi_order` oo
    LEFT JOIN panda.`aci_user_info` aui
    ON oo.`user_id` = aui.`user_id`
    WHERE oo.`create_at` > '2017-8-1 00:00:00'
    {}
    GROUP BY DATE(oo.`create_at`)
    '''
    src_cur.execute(numsql.format(from_dict[d]))
    nresult = src_cur.fetchall()
    src_cur.execute(sumsql.format(from_dict[d]))
    sresult = src_cur.fetchall()
    dst_args = []
    for n, s in zip(nresult, sresult):
        date = n[0]
        sum = s[1]
        num = n[1]
        cancel = s[1] - n[1]
        total = n[2]
        shopfrom = d
        dst_args.append((date, sum, num, cancel, total, shopfrom))

    insertsql = '''
    insert into ods.ods_month_info values (%s,%s,%s,%s,%s,%s)'''
    dst_cur.executemany(insertsql, dst_args)
    dst_con.commit()

src_cur.close()
src_con.close()
dst_cur.close()
dst_con.close()



