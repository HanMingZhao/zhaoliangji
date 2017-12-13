import requests
import xlwt
import time
import config as cf
import pymysql as db
import datetime
start = time.time()

trans = {'iPhone SE': 114,
         'iPhone5s': 5,
         'iPhone6': 4,
         'iPhone6 Plus': 16,
         'iPhone6s': 14,
         'iPhone6s Plus': 15,
         'iPhone7': 127,
         'iPhone7 Plus': 129,
         'iPhone8': 525}

wb = xlwt.Workbook()
sheet = wb.add_sheet('sheet')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '颜色')
sheet.write(0, 3, '成色')
sheet.write(0, 4, '网络')
sheet.write(0, 5, '价格')
sheet.write(0, 6, '数量')
redpattern = xlwt.Pattern()
redpattern.pattern = xlwt.Pattern.SOLID_PATTERN
redpattern.pattern_fore_colour = 2
stylered = xlwt.XFStyle()
stylered.pattern = redpattern

yellowpattern = xlwt.Pattern()
yellowpattern.pattern = xlwt.Pattern.SOLID_PATTERN
yellowpattern.pattern_fore_colour = 5
styleyellow = xlwt.XFStyle()
styleyellow.pattern = yellowpattern


respData = 'respData'
datas = 'datas'
subTitle = 'subTitle'
title = 'title'
salePrice = 'salePrice'
count = 'count'
s = ' '
ratestr = '9成新'
address = 'http://youpin.m.58.com/zzyp/list/page?filtrate=pve_5461_101_pve_5462_2101018&page={}'
products = {}
for i in range(1, 100):
    print('scanning page {}...'.format(i), time.time()-start)
    req = requests.get(url=address.format(i))
    resp = req.json()
    if resp[respData][datas] is not None:
        rows = len(sheet.rows)
        for j, r in enumerate(resp[respData][datas]):
            row = rows + j
            version = r[title].split(s)[0]
            memory = r[title].split(s)[-2]
            color = r[title].split(s)[-1]
            rate = r[subTitle].split(s)[0]
            net = r[subTitle].split(s)[1]
            price = r[salePrice]
            if len(r[title].split(s)) > 3:
                version = version + ' ' + r[title].split(s)[1]
            sheet.write(row, 0, version)
            sheet.write(row, 1, memory)
            sheet.write(row, 2, color)
            sheet.write(row, 3, rate)
            sheet.write(row, 4, net)
            sheet.write(row, 5, price)
            sheet.write(row, 6, r[count])

            if ratestr in rate:
                name = version + ':' + memory + ':' + rate
                if name in products:
                    if price > products[name]:
                        products[name] = price
                else:
                    products[name] = price
    else:
        break

option = cf.product
scon = db.connect(host=option['host'], user=option['user'], passwd=option['pass'], port=option['port'], db=option['db'])
scur = scon.cursor()
propsql = '''
SELECT ppv.pvid,ppv.pv_name FROM panda.`pdi_prop_value` ppv
WHERE ppv.pnid = {}
'''

md = {}
scur.execute(propsql.format(11))
memories = scur.fetchall()
for m in memories:
    md[str(m[0])] = m[1]

avgSql = '''
SELECT AVG(pp.`price`) FROM panda.`pdi_product` pp
LEFT JOIN panda.`pdi_model` pm
ON pp.`model_id` = pm.`model_id` 
WHERE pp.`status` = 1
AND pp.`key_props` LIKE '%12:26;%'
AND pp.`key_props` LIKE '%5:{};%'
AND pp.`key_props` LIKE '%11:{};%'
'''

sheet = wb.add_sheet('价格对比')
sheet.write(0, 0, '型号')
sheet.write(0, 1, '内存')
sheet.write(0, 2, '转转价格')
sheet.write(0, 3, '自营价格')
for tr in trans:
    for m in md:
        number = scur.execute(avgSql.format(trans[tr], m))
        avg = scur.fetchone()[0]
        if avg is not None:
            name = tr + ':' + md[m] + ':' + ratestr
            row = len(sheet.rows)
            if name in products:
                if avg > products[name]:
                    print(tr, md[m], products[name], avg, 'red')
                    print(row)
                    sheet.write(row, 0, tr, stylered)
                    sheet.write(row, 1, md[m], stylered)
                    sheet.write(row, 2, products[name], stylered)
                    sheet.write(row, 3, avg, stylered)
                    print(len(sheet.rows))
                elif avg > products[name]-50:
                    print(tr, md[m], products[name], avg, 'yellow')
                    print(row)
                    sheet.write(row, 0, tr, styleyellow)
                    sheet.write(row, 1, md[m], styleyellow)
                    sheet.write(row, 2, products[name], styleyellow)
                    sheet.write(row, 3, avg, styleyellow)
                    print(len(sheet.rows))
                else:
                    print(tr, md[m], products[name], avg, 'white')
                    print(row)
                    sheet.write(row, 0, tr)
                    sheet.write(row, 1, md[m])
                    sheet.write(row, 2, products[name])
                    sheet.write(row, 3, avg)
                    print(len(sheet.rows))

print('overtime...', time.time())
wb.save(cf.path+datetime.datetime.today().strftime(cf.date_format)+'zhuanzhuan.xls')
scur.close()
scon.close()
