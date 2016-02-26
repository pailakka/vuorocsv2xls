import xlsxwriter
import sys
import unicodecsv
import os
import urllib2
import time
import datetime
output_dir = os.environ.get('VUORO2XLS_OUTPUT','.')
TMP_PATH = os.environ.get('VUORO2XLS_TMP','.')


url = 'https://koontikartta.navici.com/tiedostot/vuoro.csv'
if len(sys.argv) == 2:
    url = sys.argv[1]
outfn = os.path.join(output_dir,'vuoro.xlsx')
fname = 'vuoro.csv'
loadfn = os.path.join(TMP_PATH,fname)

if not os.path.exists(loadfn):
    print 'loading',loadfn
    if os.path.exists(loadfn):
        os.unlink(loadfn)
    r = urllib2.urlopen(url)
    f = open(loadfn,'wb')
    f.write(r.read())
    f.close()
    r.close()
    print 'done!'


st = time.time()
wb = workbook = xlsxwriter.Workbook(outfn)

bf = wb.add_format()
bf.set_bold()

ws = wb.add_worksheet()
date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
with open(loadfn,'rb') as f:
    csvf = unicodecsv.reader(f, delimiter=';', quotechar='"',encoding='utf-8-sig')
    header = False
    r = 0
    for l in csvf:
        if not header:
            header = l

        for c,d in enumerate(l):
            if r == 0:
                ws.write(r,c,d.replace('"',''),bf)
            elif header[c].endswith('pvm') and r > 0 and len(d) > 12:
                dt = datetime.datetime.strptime(d,'%d.%m.%Y %H:%M:%S')
                ws.write_datetime(r, c, dt,date_format )
            else:
                ws.write(r,c,d)

        if r % 1000 == 0:
            print r
        r+=1


ws.autofilter(0,0,r,c-1)
ws.set_zoom(90)
wb.close()
print outfn,'done','took',time.time()-st
