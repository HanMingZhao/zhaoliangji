import pymysql as db
import numpy as np

src_con = db.connect(host='rm-bp13wnvyc2dh86ju1.mysql.rds.aliyuncs.com', user='panda_reader', passwd='zhaoliangji3503',
                     db='panda', charset='utf8')
dst_con = db.connect(host='114.215.176.190', user='root', passwd='huodao123', db='ods', port=33069, charset='utf8')

src_cur = src_con.cursor()
dst_cur = dst_con.cursor()

times = ["afs.times < 1           ",
         "afs.times < 2 and afs.times > 1  ",
         "afs.times < 3 and afs.times > 2  ",
         "afs.times < 4 and afs.times > 3  ",
         "afs.times < 5 and afs.times > 4  ",
         "afs.times < 6 and afs.times > 5  ",
         "afs.times < 7 and afs.times > 6  ",
         "afs.times < 8 and afs.times > 7  ",
         "afs.times < 9 and afs.times > 8  ",
         "afs.times < 10 and afs.times > 9 ",
         "afs.times < 11 and afs.times > 10 ",
         "afs.times < 12 and afs.times > 11 ",
         "afs.times < 13 and afs.times > 12 ",
         "afs.times < 14 and afs.times > 13 ",
         "afs.times < 15 and afs.times > 14 "]

af_types = [1, 2, 3]

for af_type in af_types:
    countsql = '''
          SELECT DATE(ooa.`sendback_time`),COUNT(1) FROM panda.`odi_order_aftersale` ooa
          WHERE ooa.`type` = {} 
          AND ooa.`sendback_time` > '2017-09-12 00:00:00'
          AND ooa.`sendback_time` < DATE(NOW())
          GROUP BY DATE(ooa.`sendback_time`)
          '''
    src_cur.execute(countsql.format(af_type))
    af_result = src_cur.fetchall()
    af_array = np.zeros((len(af_result), len(times)), dtype=int)
    afl = af_array.tolist()
    for x, y in zip(af_result, afl):
        y.insert(0, x[1])
        y.insert(0, str(x[0]))
    for i, n in enumerate(times):
        timecountsql = '''
               SELECT afs.dates,COUNT(1) FROM 
               (
               SELECT DATE(ooa.`sendback_time`) `dates`,((UNIX_TIMESTAMP(ooa.`finsh_time`)-UNIX_TIMESTAMP(ooa.`sendback_time`))/60/60)/24 `times` FROM panda.`odi_order_aftersale` ooa
               WHERE ooa.`type` = {}
               AND (UNIX_TIMESTAMP(ooa.`finsh_time`)-UNIX_TIMESTAMP(ooa.`sendback_time`))>0
               AND ooa.`sendback_time` >'2017-09-12 00:00:00'
               AND ooa.`sendback_time` <DATE(NOW())
               GROUP BY DATE(ooa.`sendback_time`),`times`
               ) afs
               WHERE {}
               GROUP BY afs.dates
              '''
        src_cur.execute(timecountsql.format(af_type, n))
        af_times = src_cur.fetchall()
        for z in afl:
            for j in af_times:
                if str(j[0]) == z[0]:
                    z[i+2] = j[1]
                else:
                    continue

    dst_args = []

    tablesql = '''drop table if EXISTS ods.ods_aftersale_{} ;
        create TABLE if NOT EXISTS ods.ods_aftersale_{}
        ( `日期` VARCHAR(32),
        `收到` VARCHAR(4),
        `1天内` VARCHAR(4),
        `2天内` VARCHAR(4),
        `3天内` VARCHAR(4),
        `4天内` VARCHAR(4),
        `5天内` VARCHAR(4),
        `6天内` VARCHAR(4),
        `7天内` VARCHAR(4),
        `8天内` VARCHAR(4),
        `9天内` VARCHAR(4),
        `10天内` VARCHAR(4),
        `11天内` VARCHAR(4),
        `12天内` VARCHAR(4),
        `13天内` VARCHAR(4),
        `14天内` VARCHAR(4),
        `15天内` VARCHAR(4)
        )
        ENGINE=MYISAM CHARSET=utf8;
        '''
    dst_cur.execute(tablesql.format(af_type, af_type))
    dst_con.commit()

    for row in afl:
        dst_args.append(tuple(str(x) for x in row))
    print(dst_args)
    insert = '''
             insert into ods.ods_aftersale_{}  values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
             '''
    dst_cur.executemany(insert.format(af_type), dst_args)
    dst_con.commit()

src_cur.close()
src_con.close()
dst_cur.close()
dst_con.close()
