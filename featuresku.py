import pymysql as db
import xlwt
import numpy as np
import datetime as dt
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

dst_host = cf.get('test', 'host')
dst_user = cf.get('test', 'user')
dst_pass = cf.get('test', 'passwd')
dst_port = cf.getint('test', 'port')
dst_db = cf.get('test', 'db')
dcon = db.connect(host=dst_host, user=dst_user, passwd=dst_pass, db=dst_db, port=dst_port, charset='utf8')
dcur = dcon.cursor()

warehouse = {1: '分拾', 2: '检测', 3: '市场', 4: '上架', 5: '维修', 6: '报废', 7: 'B端', 8: '预上架', 9: '外包维修',
             11: '京东', 12: '待卖'}

for wnum in warehouse:

    pass

dcur.close()
dcon.close()
scur.close()
scon.close()
