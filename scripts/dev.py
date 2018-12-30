import datetime
import os 
import sys
import logging
import win32com.client
import csv
os.umask(0000)
os.path.splitext(os.path.basename(__file__))[0]
strpath=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
######################################################
#                   Set Config Path
######################################################
strconfigdir=strpath+'\\config\\PyRunExcel_config.csv'
if  len(sys.argv)>1 :
    strconfigdir=sys.argv[0]
logging.basicConfig(filename=strpath+'\\log\\' \
                    +os.path.splitext(os.path.basename(__file__))[0]+'_' \
                    +datetime.datetime.today().strftime('%Y%m%d')+'.txt',level=logging.DEBUG \
                    ,format='%(asctime)s.%(msecs)03d %(levelname)s %(module)s - %(funcName)s: %(message)s' \
                    ,datefmt="%Y-%m-%d %H:%M:%S")
logging.info('Start ')
iRow=0
oConfig = open(strconfigdir, 'rb')
reader = csv.DictReader(oConfig)
xlApp=win32com.client.DispatchEx("Excel.Application")
for row in reader:
    if row['ExcelPath']!='':
        xlApp.Workbooks.Open(Filename=row['ExcelPath']) if row['isReadOnly']=='TRUE' \
        else xlApp.Workbooks.Open(Filename=row['ExcelPath'], ReadOnly=True,Password='')
    if row['Parameters']=='':
        exec("xlApp.Application.Run(row['RunProcess'])") 
    else :
        exec("xlApp.Application.Run(row['RunProcess']"+row['Parameters']+")")
    if row['isContinuous']=='FALSE':
        xlApp.ActiveWorkbook.Close(SaveChanges=True) if row['isSave']=='TRUE' \
        else xlApp.ActiveWorkbook.Close(SaveChanges=False) 
        xlApp.Quit()
        xlApp = None 
    iRow=iRow+1
del xlApp 
oConfig.close()

