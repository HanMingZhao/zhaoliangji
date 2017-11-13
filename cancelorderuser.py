import pymysql as db
import xlwt
import configparser

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

style = xlwt.XFStyle()
style.num_format_str = 'YYYY-MM-DD h:mm:ss'
# Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0

for i in range(6):
    querysql = '''
    SELECT tbt.user_id `用户id`,tbt.count `下单次数`,tbt.create_at `最后下单时间`, tbt.contacts `联系人`,tbt.phone `联系电话` FROM 
    (
    SELECT bt.user_id,COUNT(1) `count`,bt.`create_at`,bt.`product_id`,bt.total_amount,bt.contacts,bt.phone  FROM 
    (
    SELECT oo.`user_id`,COUNT(1) `count`,oo.`create_at`,oo.`product_id`,oo.`total_amount`,oo.`contacts`,oo.`phone` FROM panda.`odi_order` oo
    WHERE oo.`order_status` = 3
    AND oo.`user_id` NOT IN 
    (SELECT DISTINCT(oo.`user_id`) FROM panda.`odi_order` oo
    WHERE oo.`order_status` IN (1,2,4,5))
    AND oo.`create_at`>'2017-10-1'
    GROUP BY oo.`user_id`,oo.`product_id`,oo.`total_amount`#DAY(oo.create_at),HOUR(oo.create_at)#,MINUTE(oo.create_at)#,second(oo.`create_at`)
    ORDER BY oo.`create_at` DESC
    )bt
    GROUP BY bt.user_id
    )tbt
    WHERE tbt.count>3
    limit {},213
    '''
    page = i + 1
    limit = 213 * (page - 1) + 1
    scur.execute(querysql.format(limit))
    result = scur.fetchall()
    sheet = wb.add_sheet(str(page))
    sheet.write(0, 0, '用户id')
    sheet.write(0, 1, '下单次数')
    sheet.write(0, 2, '最后下单时间')
    sheet.write(0, 3, '联系人')
    sheet.write(0, 4, '联系电话')
    for j, r in enumerate(result):
        sheet.write(j+1, 0, r[0])
        sheet.write(j+1, 1, r[1])
        sheet.write(j+1, 2, r[2], style)
        sheet.write(j+1, 3, r[3])
        sheet.write(j+1, 4, r[4])

path = cf.get('path', 'path')
wb.save(path + 'cancelorder.xls')

scur.close()
scon.close()
