import pymysql as db
import config as conf
import xlwt

cf = conf.product
connect = db.connect(host=cf['host'], user=cf['user'], passwd=cf['pass'], port=cf['port'], charset=conf.char)
cursor = connect.cursor()
wb = xlwt.Workbook()

cursor.close()
connect.close()
