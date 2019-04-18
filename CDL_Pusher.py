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

    def move_files(self,batches,test = True):
        if test != True:
            write_path = 'Z:\\'
        else:
            write_path = 'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\CLD_PUSH_TEST\\'
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

def push_dataniche(qtr,year, test = True):
    path = r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\DataNicheTest2\Claims'
    to_submit = []
    for root, dirs, files in os.walk(path):
        for file in files:
            if f'{qtr}Q{year}'in file:
                file_name = os.path.join(root,file)
                to_submit.append(file_name)
    n = 40
    print(f'Number of files to submit for OV3 is {len(to_submit)}')
    batches = [to_submit[i:i+n] for i in range(0,len(to_submit),n)]
    if test != True:
        write_path = 'Z:\\'
    else:
        write_path = 'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\CLD_PUSH_TEST\\'
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
    batches = pusher.batch_files(qtr='4',year='2018')
    pusher.move_files(batches, test=False)
    #push_dataniche('4','2018', test=False)

if __name__ == '__main__':
    main()
    

