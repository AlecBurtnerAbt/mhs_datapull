# -*- coding: utf-8 -*-
"""
Created on Mon Oct 15 12:52:42 2018

@author: C252059
"""

import os
import pandas as pd
import re
import shutil
import time

class Pusher():
    
    def __init__(self):
        self.path = 'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\'
        os.chdir(self.path)
    
    def batch_files(self,qtr,year):
        to_submit = []
        for root, dirs, files in os.walk(self.path):
            
            for file in files:
                if 'sup'in file.lower():
                    pass
                elif qtr in root and year in root:
                    file_name = os.path.join(root,file)
                    to_submit.append(file_name)
                '''
                conv = fr'^\w{{2}}[_].+[_{qtr}Q{year}].*[.]\w{{3}}'
                if re.match(conv,file):
                    #to_submit.append(file_name)
                    print(file+' Matches!')
                    '''
        n = 40
        batches = [to_submit[i:i+n] for i in range(0,len(to_submit),n)]
        return batches

    def move_files(self,batches):
        write_path = 'Z:\\'
        for batch in batches:
            for file in batch:
                base_file = file.split('\\')[-1].replace(' ','-')
                extension = base_file.split('.')[-1]
                base_file = base_file.split('.')[0]
                base_file = base_file+"_LABNET_"
                print(f"base file is {base_file}")
                print(f"extension is {extension}")
                base_file = base_file+'.'+extension
                print(f"base file is now {base_file}")
                file_name = 'IRIS.CLD.'+base_file
                print(file_name)
                if file_name[-4:]=='xlsx' or file_name[-4:]=='.xls':
                    shutil.copy(file,write_path+file_name)
                    print(f'{file_name} has been moved to the magic folder!  Pray to your god it works!')
                else:
                    pass
            while len(os.listdir(write_path))>1:
                time.sleep(1)
    
    
def main():
    pusher = Pusher()
    batches = pusher.batch_files(qtr='3',year='2018')
    pusher.move_files(batches)

if __name__ == '__main__':
    main()
    
    
'''
The function below is a block of code if you have to push one file
from "batches" to a folder to test file transfer functionality

'''
'''
def one_off_push(file):
    write_path = 'Z:\\'
    base_file = file.split('\\')[-1].replace(' ','-')
    extension = base_file.split('.')[-1]
    base_file = base_file.split('.')[0]
    base_file = base_file+"_LABNET_"
    print(f"base file is {base_file}")
    print(f"extension is {extension}")
    base_file = base_file+'.'+extension
    print(f"base file is now {base_file}")
    file_name = 'IRIS.CLD.'+base_file
    print(file_name)
    shutil.copy(file,write_path+file_name)

file = batches[0][0]
one_off_push(file)
'''
