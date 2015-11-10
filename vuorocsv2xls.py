import xlsxwriter
import sys
import unicodecsv
import os
import wget
import time
output_dir = os.environ.get('VUORO2XLS_OUTPUT','.')
url = 'https://koontikartta.navici.com/tiedostot/vuoro.csv'
if len(sys.argv) == 2:
    url = sys.argv[1]

fname = 'vuoro.csv'

loadfn = os.path.join(output_dir,fname)
outfn = os.path.join(output_dir,'vuoro.xlsx')
print 'loading',loadfn
if os.path.exists(loadfn):
    os.unlink(loadfn)
wget.download(url,loadfn)
print


st = time.time()
wb = workbook = xlsxwriter.Workbook(outfn)

bf = wb.add_format()
bf.set_bold()

ws = wb.add_worksheet()

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
            else:
                ws.write(r,c,d)

        if r % 1000 == 0:
            print r
        r+=1


ws.autofilter(0,0,r,c-1)
ws.set_zoom(90)
wb.close()
print outfn,'done','took',time.time()-st