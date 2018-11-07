# -*- coding: utf-8 -*-
"""
Created on Mon Aug 13 16:06:40 2018

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

class DelawareGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Delaware')
        
        
    def pull(self):
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Delaware', usecols=[0,1],dtype='str')
        login_credentials = self.credentials
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        #Get the program map.  This map must be maintainted by the MHS team.
        
        mapper = dict(zip(mapper.Delaware,mapper.Lilly))
        #build date time group
        yq = str(yr)+'q'+str(qtr)
        #Get to the page and login
        wait = WebDriverWait(driver,10)
        wait2 = WebDriverWait(driver,2)
        driver.get('https://www.edsdocumenttransfer.com/')
        user = wait.until(EC.element_to_be_clickable((By.ID,'form_username')))
        user.send_keys('llymedicaid@lilly.com')
        password = driver.find_element_by_id('form_password')
        password.send_keys('Spring16!')
        login = driver.find_element_by_id('submit_button')
        login.click()
        
        #Wait until the folder dropdown is available then 
        #select the distribution folder
        folders = wait.until(EC.element_to_be_clickable((By.ID,'field_gotofolder')))
        folders_select = Select(folders)
        folders_select.select_by_visible_text('/ Distribution')
        #now give it some time to load
        time.sleep(2)
        sub_folders = lambda: driver.find_elements_by_xpath('//table[@id="folderfilelisttable"]//tr//td//img[@title="Folder"]')
        invoices = []
        for k,folder in enumerate(sub_folders()):
            sub_folders()[k].click()
            wait.until(EC.element_to_be_clickable((By.XPATH,'//span[text()="Parent Folder"]')))
            try:
                num_pages = driver.find_element_by_xpath('//table//tbody//tr[@class="nullSpacer"]//td//b')
                num_pages = num_pages.text[-1:]
            except NoSuchElementException as ex:
                num_pages = 1
            for i in range(int(num_pages)):
                new_files = lambda: driver.find_elements_by_xpath('//img/following-sibling::span[contains(text(),"%s")]'%(yq))
                if i==0:
                    pass
                else:
                    nex = driver.find_element_by_xpath('//span[text()="Next"]')
                    nex.click()
                    wait.until(EC.element_to_be_clickable((By.XPATH,'//span[text()="Parent Folder"]')))
                for j, file in enumerate(new_files()):
                    name = new_files()[j].text
                    Name = name[6:]
                    label_code = Name[:5]
                    file_type = Name[5:9]
                    program = name[name.find('q')+2:name.find('.')].lower()
                    ext = Name[-3:]
                    lilly_program = mapper[program]
                    new_files()[j].click()
                    file_name = lilly_program+'_'+label_code+'_'+str(yr)+'_'+str(qtr)
                    if file_type =='clda':
                        download = wait2.until(EC.element_to_be_clickable((By.ID,'downloadLink')))
                        download.click() 
                        ext = '.txt'
                        file_name = f'DE_{lilly_program}_{qtr}Q{yr}_{label_code}{ext}'
                        try:
                            close_pop_up = wait.until(EC.element_to_be_clickable((By.XPATH,'//div//ips-verifier//div//div//div//i')))
                            close_pop_up.click()
                        except TimeoutException as ex: 
                            pass
                        while name not in os.listdir():
                            time.sleep(1)
                    elif file_type=='invd' and ext =='dat':
                        download = wait2.until(EC.element_to_be_clickable((By.ID,'downloadLink')))
                        download.click() 
                        ext = '.txt'
                        file_name = file_name+ext
                        try:
                            close_pop_up = wait.until(EC.element_to_be_clickable((By.XPATH,'//div//ips-verifier//div//div//div//i')))
                            close_pop_up.click()
                        except TimeoutException as ex: 
                            pass
                        while name not in os.listdir():
                            time.sleep(1)
                    else:
                        download = wait.until(EC.element_to_be_clickable((By.XPATH,'//a//span[text()="Download"]')))
                        download.click()
                        ext = '.pdf'
                        file_name = file_name+ext
                        invoices.append(file_name)
                        try:
                            close_pop_up = wait.until(EC.element_to_be_clickable((By.XPATH,'//div//ips-verifier//div//div//div//i')))
                            close_pop_up.click()
                        except TimeoutException as ex: 
                            pass                    
                        while name not in os.listdir():
                            time.sleep(1)
                        
                    if file_type == 'clda':
                        file_type = 'Claims'
                    else:
                        file_type='Invoices'
                    
                    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+file_type+'\\'+'Delaware'+'\\'+lilly_program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                    
                    if os.path.exists(path)==False:
                        os.makedirs(path)
                    else:
                        pass
        
                    shutil.move(name,path+file_name)
        
                    driver.back()
                    time.sleep(1.5)
            folders = wait.until(EC.element_to_be_clickable((By.ID,'field_gotofolder')))
            folders_select = Select(folders)
            folders_select.select_by_visible_text('/ Distribution')
        q_flag = 0
        while q_flag==0:
            try:
                self.driver.close()
                q_flag=1
            except MaxRetryError as err:
                continue
        os.chdir('O:\\')
        os.removedirs(self.temp_folder_path)
        return(invoices)
        
    def morph_invoices(self):
        path = f'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Delaware\\'
        for root, folders, files in os.walk(path):
                for file in files:
                    if file[-3:]=='txt':
                        with open(root+'\\'+file) as f:
                            file_data = []
                            lines = f.readlines()
                            for line in lines:
                                CDE_NDC = line[:11]
                                CDE_CLM_TYPE = line[11]
                                CDE_ICN = line[12:25]
                                IND_ADJ = line[25]
                                NUM_DTL = line[26:30]
                                CDE_PROC = line[30:36]
                                ID_PROVIDER = line[36:51]
                                CDE_FUND_CODE = line[51:55]
                                DTE_FDOS = line[55:63]
                                DTE_PAID = line[63:71]
                                NUM_DAYS_SUPPLY = line[71:75]
                                QTY_UNITS_BILLED = line[75:89]
                                AMT_BILLED = line[89:100]
                                AMT_PD_MCAID = line[100:111]
                                AMT_PD_NON_MCAID = line[111:122]
                                IND_TPL = line[122]
                                AMT_ALWD = line[123:134]
                                DTE_ADJUDICATED = line[134:]
                                file_data.append([CDE_NDC,CDE_CLM_TYPE,CDE_ICN,IND_ADJ,NUM_DTL,CDE_PROC,ID_PROVIDER,CDE_FUND_CODE,DTE_FDOS,
                                                  DTE_PAID,NUM_DAYS_SUPPLY,QTY_UNITS_BILLED,AMT_BILLED,AMT_PD_MCAID,AMT_PD_NON_MCAID,
                                                  IND_TPL,AMT_ALWD,DTE_ADJUDICATED])
                        file_name = file[:-4]+'.xlsx'
                        df = pd.DataFrame(file_data,columns=['CDE_NDC','CDE_CLM_TYPE','CDE_ICN','IND_ADJ','NUM_DTL','CDE_PROC','ID_PROVIDER','CDE_FUND_CODE','DTE_FDOS',
                                          'DTE_PAID','NUM_DAYS_SUPPLY','QTY_UNITS_BILLED','AMT_BILLED','AMT_PD_MCAID','AMT_PD_NON_MCAID',
                                          'IND_TPL','AMT_ALWD','DTE_ADJUDICATED'])
                        placement_path = root+'\\'+file_name
                        df.to_excel(placement_path,index=False)
                        os.remove(root+'\\'+file)        
def main():
    grabber = DelawareGrabloid()
    invoices = grabber.pull()
    grabber.send_message(invoices)
    grabber.morph_invoices()
if __name__ =='__main__':
    main()
