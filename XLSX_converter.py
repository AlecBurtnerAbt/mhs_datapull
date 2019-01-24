# -*- coding: utf-8 -*-
"""
Created on Thu Oct 18 13:34:11 2018

@author: C252059
"""

import os
import win32com.client as com
import shutil
qtr = '3'
yr = '2018'
path = r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\Claims'

#The below loop goes through and changes all XLS files to XLSX files
def convert(path):
    for root, dirs, files in os.walk(path):
        for file in files:
            if f'{qtr}Q{yr}' not in file:
                continue
            if file.split('.')[-1]=='xls':
                excel = com.gencache.EnsureDispatch('Excel.Application')
                wb = excel.Workbooks.Open(root+'\\'+file)
                wb.SaveAs(root+'\\'+file+'x', FileFormat=51)
                wb.Close()
#The below file moves all XLS files from the Test\Claims directory to the archive
def archive():
    path2 = r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\Archive\XLS'       
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.split('.')[-1]=='xls':
                shutil.move(root+'\\'+file,path2+'\\'+file)
   
def list_files():
    import pandas as pd  
    to_submit = []       
    for root, dirs, files in os.walk(path):
        for file in files:
            if f'{qtr}Q{yr}' not in file:
                continue
            to_submit.append(file)
    to_submit = pd.DataFrame(to_submit,columns=['Files'])
    to_submit.to_excel(os.path.join(path,'submitted_files.xlsx'),engine='xlsxwriter',index=False)

if __name__=="__main__":
    convert(path)
    archive()
    list_files()
