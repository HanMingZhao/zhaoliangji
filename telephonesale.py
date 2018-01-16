import pandas as pd
import config as conf
import xlwt

dframe = pd.read_csv('13-16.csv')
wb = xlwt.Workbook()
two_day_sale_sql = '''
SELECT * FROM panda.`odi_order` oo
WHERE oo.`order_status` IN (1,2,4,5)
AND oo.`order_type` IN (1,2)
AND oo.`pay_at`>'{}'
AND oo.`pay_at`<'{}'
AND oo.`phone`= '{}'
'''
saled = []
for phone in dframe['phone']:
    count = conf.product_cursor.execute(two_day_sale_sql.format(conf.yesterday.strftime(conf.date_format),
                                                        conf.tomorrow.strftime(conf.date_format), phone))
    print(phone)
    if count > 0:
        saled.append(phone)
sheet = wb.add_sheet('sheet')
for i, s in enumerate(saled):
    sheet.write(i, 0, str(s))
wb.save(conf.today.strftime(conf.date_format) + 'telesale.xls')
conf.product_cursor.close()
conf.product_connect.close()
