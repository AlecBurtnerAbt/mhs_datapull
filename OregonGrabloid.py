# -*- coding: utf-8 -*-
"""
Created on Tue Jun  4 10:21:21 2019

@author: AUTOBOTS
"""

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
from grabloid import Grabloid, push_note

class OregonGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Oregon')
        
        
    def pull(self):
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Oregon', usecols='E,F',dtype='str')
        modifier_table =  pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Oregon', usecols='C,D',dtype='str')
        mapper = dict(zip(mapper.Portal,mapper.Flex))
        modifier_table = dict(zip(modifier_table.Modifier,modifier_table.Meaning))
        login_credentials = self.credentials
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]

    
        #build date time group
        yq = str(yr)+'q'+str(qtr)
        #Get to the page and login
        wait = WebDriverWait(driver,10)
        wait2 = WebDriverWait(driver,2)
        driver.get('https://www.edsdocumenttransfer.com/')
        user = wait.until(EC.element_to_be_clickable((By.ID,'form_username')))
        user.send_keys(username)
        password_input = driver.find_element_by_id('form_password')
        password_input.send_keys(password)
        login = driver.find_element_by_id('submit_button')
        login.click()
        
        #Wait until the folder dropdown is available then 
        #select the distribution folder
        folders = wait.until(EC.element_to_be_clickable((By.ID,'field_gotofolder')))
        folders_select = Select(folders)
        folders_select.select_by_visible_text('/ Distribution')
        #now give it some time to load
        oregon_folder = wait.until(EC.presence_of_element_located((By.XPATH,'//span[text()="Oregon"]')))
        oregon_folder.click()
        sub_folders = lambda: driver.find_elements_by_xpath('//table[@id="folderfilelisttable"]//tr//td//img[@title="Folder"]')
        invoices = []
        for k,folder in enumerate(sub_folders()):
            sub_folders()[k].click()
            current_quarter_folder = wait.until(EC.presence_of_element_located((By.XPATH,f'//span[text()="{"".join([str(yr),str(qtr)])}"]')))
            current_quarter_folder.click()
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
                    labeler = name[:5]
                    program_code = name.split('.')[0][-3:-1]
                    modifier = name.split('.')[0][-1]
                    flex_code = mapper[program_code]
                    modifier = modifier_table[modifier]
                    program = name[name.find('q')+2:name.find('.')].lower()
                    extension = name[-3:]
                    new_files()[j].click()
                    file_name = f'OR_{flex_code}_{qtr}Q{yr}_{labeler}_{modifier}_.{extension}'
                    file_type = name[5:9]
                    if file_type =='clda':
                        download = wait2.until(EC.element_to_be_clickable((By.ID,'downloadLink')))
                        download.click() 
                        ext = '.txt'
                        download_file_name = f'OR_{flex_code}_{qtr}Q{yr}_{labeler}_{modifier}_{ext}'
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
                    
                    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+file_type+'\\'+'Oregon'+'\\'+flex_code+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                    
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
            oregon_folder = wait.until(EC.presence_of_element_located((By.XPATH,'//span[text()="Oregon"]')))
            oregon_folder.click()
        q_flag = 0
        while q_flag==0:
            try:
                self.driver.close()
                q_flag=1
            except MaxRetryError as err:
                continue
        return(invoices)
        
    def morph_invoices(self):
        path = f'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Oregon\\'
        print('Converting CLD .dat to .xlsx')
        for root, folders, files in os.walk(path):
                for file in files:
                    if file[-3:]=='dat':
                        print(f'Working on {file}')
                        with open(root+'\\'+file) as f:
                            file_data = []
                            lines = f.readlines()
                            for line in lines:
                                ICN = line[:13]
                                VOID_ICN = line[13:26]
                                CDE_NDC = line[26:37]
                                RBT_PERIOD = line[37:42]
                                QTY_REBATE = line[42:58]
                                DTE_SERVICE = line[58:66]
                                DTE_PAID = line[66:74]
                                ID_PROV_BILLING = line[74:89]                                
                                ID_PROV_PRESCRB = line[89:104]
                                AMT_PAID = line[104:115]                                
                                AMT_BILLED = line[115:126]                                
                                AMT_COPAY = line[126:133]                                
                                AMT_TPL = line[133:146]
                                AMD_NDCPRO = line[146:157]
                                NUM_PRESCRIP = line[157:169]
                                NUM_DAY_SUPPLY = line[169:173]
                                QTY_REFILL = line[173:]                                
                                file_data.append([ICN, VOID_ICN,CDE_NDC,RBT_PERIOD,QTY_REBATE,DTE_SERVICE,DTE_PAID,ID_PROV_BILLING,
                                                  ID_PROV_PRESCRB,AMT_PAID,AMT_BILLED,AMT_COPAY,AMT_TPL,AMD_NDCPRO,NUM_PRESCRIP,
                                                  NUM_DAY_SUPPLY,QTY_REFILL])
                        file_name = file[:-4]+'.xlsx'
                        df = pd.DataFrame(file_data,columns=['ICN', 'VOID_ICN','CDE_NDC','RBT_PERIOD','QTY_REBATE'
                                                              ,'DTE_SERVICE','DTE_PAID','ID_PROV_BILLING',
                                                  'ID_PROV_PRESCRB','AMT_PAID','AMT_BILLED','AMT_COPAY','AMT_TPL'
                                                  ,'AMD_NDCPRO','NUM_PRESCRIP','NUM_DAY_SUPPLY','QTY_REFILL'])
                        placement_path = root+'\\'+file_name
                        df.to_excel(placement_path,index=False)
                        os.remove(root+'\\'+file)        
                        
#@push_note(__file__)
def main():
    grabber = OregonGrabloid()
    invoices = grabber.pull()
    grabber.morph_invoices()
    grabber.cleanup()
    grabber.send_message(invoices)
    
if __name__ =='__main__':
    main()


