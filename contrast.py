import config as cf
import datetime as dt
import xlwt
import pymysql as db

option = cf.test
src_con = db.connect(host=option['host'], user=option['user'], passwd=option['pass'], port=option['port'],
                     db=option['db'])
src_cur = src_con.cursor()

wb = xlwt.Workbook()
sheet = wb.add_sheet('sheet')




src_cur.close()
src_con.close()
