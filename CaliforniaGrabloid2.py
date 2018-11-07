# -*- coding: utf-8 -*-
"""
Created on Fri Jul 20 08:14:27 2018

@author: C252059
"""
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import time
import os
from win32com.client import Dispatch
import pandas as pd
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from bs4 import BeautifulSoup
import gzip
import shutil
import zipfile
import pandas as pd
import itertools    
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
import pprint
import gzip
import numpy as np
import xlsxwriter as xl
from grabloid import Grabloid

class CaliforniaGrabloid2(Grabloid):
    def __init__(self):
        super().__init__(script='California')
        script = self.script
        self.mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='{}'.format(self.script), usecols='E:G',dtype='str')
    
    def pull(self):
        driver = self.driver
        yr = self.yr
        qtr = self.qtr
        program_mapper = self.mapper
        login_credentials = self.credentials
        wait = self.wait
        cld_obtained = []
        mapper = dict(zip(program_mapper['Code on CA Invoice'],program_mapper['Contract ID in MRB']))
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        #navigate to the drug rebate invoice page
        user_name2 = login_credentials.iloc[1,0]
        pass_word2 = login_credentials.iloc[1,1]
        #get the three labeler codes.  Will have to update if labeler codes change
        lilly_code = lambda : wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/a')))
        dista_code = lambda : wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/a')))
        imclone_code = lambda :wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td[2]/a')))
        
        codes = [lilly_code,dista_code,imclone_code]
        '''
        This block of code downloads all of the prepared reports.  The reports come in a .gz file
        and have to be decompressed, this happens after the download in the next loop.
        '''
        for user, password in zip(login_credentials.Username[:2],login_credentials.Password[:2]):      
            driver.get('https://www.medi-cal.ca.gov/')
            transaction_tab = driver.find_element_by_xpath('//a[text()="Transactions"]')
            transaction_tab.click()
            user_name = wait.until(EC.element_to_be_clickable((By.ID,'UserID')))
            user_name.send_keys(user)
            pass_word = driver.find_element_by_id('UserPW')
            pass_word.send_keys(password)
            submit_button = driver.find_element_by_id('cmdSubmit')
            submit_button.click()             
            drug_rebate = driver.find_element_by_xpath('//*[@id="tabpanel_1_sublist"]/li/a')
            drug_rebate.click()
            
            for code in codes:
                print('getting {}'.format(code().text))
                code().click()
                retrieve = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="middle_column"]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[3]/td/a[2]/b')))
                retrieve.click()               
                wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="left_column"]/div[1]/a/img')))     
                soup2 = BeautifulSoup(driver.page_source,'html.parser') 
                bodies = soup2.find_all('tbody')
                body = bodies[2]
                rows = body.find_all('tr')
                data = body.find_all('td')
                data = [x.text for x in data]
                data = np.asarray(data)
                data = data.reshape((-1,3))
                data_df = pd.DataFrame(data)
                data_df['NDCs'] = data_df.iloc[:,2].apply((lambda x: x.split(';')[4]))
                data_df['Program'] = data_df.iloc[:,2].apply((lambda x: x.split(';')[3].split('=')[1]))
                data_df['Year'] = data_df.iloc[:,2].apply((lambda x: x.split(';')[1].split('=')[1]))
                data_df['Qtr'] = data_df.iloc[:,2].apply((lambda x: x.split(';')[2].split('=')[1]))
                data_df = data_df.drop_duplicates(subset=['NDCs','Program'])
                data_df = data_df[(data_df['Year'] ==str(yr))&(data_df['Qtr']==str(qtr))]
                data_df = data_df[data_df[1]=='\nCompleted Successfully\n']             
                links = list(data_df.iloc[:,0].apply((lambda x: x.strip())))     
                for link in range(len(links)):
                    xpath = "//a[contains(text(),'"+links[link]+"')]"
                    success=0
                    while success==0:
                        DL_link = driver.find_element_by_xpath(xpath)
                        print(f'Download {DL_link.text}')
                        DL_link.click()
                        counter = 0
                        while links[link] not in os.listdir() and counter <20:
                            time.sleep(1)
                            counter+=1
                        if any(map((lambda x: links[link] in x),os.listdir())):
                            success=1
                        if success==1:
                            pass
                        else:
                            driver.refresh()
                        
                driver.get(r'https://rais.medi-cal.ca.gov/drug/DrugLablr.asp')
            exit_link = driver.find_element_by_xpath('//a[text()="Exit"]')
            exit_link.click()
            
                
            '''
            This is the loop that goes through the downloaded .gz files, unzips them, renames them
            to the file format, makes them a text file, and then deletes the .gz file
            '''
        files = os.listdir()
        num_reports = len(files)
        for file in files:
            prog_code = file.split('_')[-1].split('.')[0]
            prog = mapper[prog_code]
            label_code = file.split('_')[2]
            request_number = file.split('_')[1]
            new_name =  'CA_{}_{}Q{}_{}_{}.txt'.format(prog,qtr,yr,label_code,request_number)
            unzipped_name = file[:-3]
            try:
                with gzip.open(file,'rt') as ref:
                    content = ref.read()
                    text_file = open(unzipped_name,'w')
                    text_file.write(content)
                    text_file.close()
                    ref.close()
            except:
                pass
            path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\California\\'+prog+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
            if os.path.exists(path)==False:
                os.makedirs(path)
            else:
                pass
            cld_obtained.append(new_name)
            shutil.move(unzipped_name,path+new_name)
            os.remove(file)
        driver.close()
        os.chdir('O:\\')
        os.removedirs(self.temp_folder_path)
        return cld_obtained
    
    def morph_cld(self):
        cali_columns =['Claim Control Number','NDC Code','Date of Service','Claim Adjudication Date','Units of Service',
          'Reimbursed Amount','Billed Amount','Adjustment Indicator','RX_ID','Billing Provider Number','Billing Provider Owner Number',
          'Billing Provider Service Location Number','Adjustment Claim Control Number','Recipient Other Coverage Code',
          'Other Health Coverage Indicator','TAR Control Number','Third Party Code','Third Party Amount','Patient Liability Amount',
          'Co-Pay Code','Co-Pay Amount','Days Supply Number','Referring Prescribing Provider Number','Recipient Crossover Status Code',
          'Recipient Health Plan Status Code','Compound Code','Cost Basis Determination Code']


        cali_compound_columns =['Claim Control Number','NDC Code','Date of Service','Claim Adjudication Date','Units of Service',
          'Reimbursed Amount','Billed Amount','Adjustment Indicator','RX_ID','Billing Provider Number','Billing Provider Owner Number',
          'Billing Provider Service Location Number','Adjustment Claim Control Number','Recipient Other Coverage Code',
          'Other Health Coverage Indicator','TAR Control Number','Third Party Code','Third Party Amount','Patient Liability Amount',
          'Co-Pay Code','Co-Pay Amount','Days Supply Number','Referring Prescribing Provider Number','Recipient Crossover Status Code',
          'Recipient Health Plan Status Code','Compound Code','Ingredient Cost Basis Determination Code', 'Claim Compound Ingredient Reimbursement Amount']
        
        path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\California\\'
        os.chdir(path)
        dirs = [dirs for roots, dirs, files in os.walk(path)]
        programs = [program for program in dirs[0]]
        for program in programs:
           path = f'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\California\\{program}\\2018\\Q2\\'
           os.chdir(path)
           if 'CMPD' in program:
               program_df = pd.DataFrame(columns= cali_compound_columns)
           else:
               program_df = pd.DataFrame(columns = cali_columns)
           for file in os.listdir():
               if file.split('.')[1] != 'txt':
                   continue
               else:
                  temp = pd.read_table(file,sep='~',dtype=str,names=program_df.columns,encoding='latin-1')
                  program_df = program_df.append(temp)
                  os.remove(file)
           new_name = path+f'CA_{program}_{self.qtr}Q{self.yr}.xlsx'
           program_df.to_excel(new_name, index=False, engine='xlsxwriter')
           
            
def main():
    grabber = CaliforniaGrabloid2()
    cld_obtained = grabber.pull()
    grabber.morph_cld()
    grabber.send_message(cld_obtained)
if __name__ == '__main__':
    main()    