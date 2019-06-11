# -*- coding: utf-8 -*-
"""
Created on Tue Aug 21 12:43:23 2018

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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoAlertPresentException
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
from pandas.errors import EmptyDataError
class Alabama_Grabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Alabama')
        self.usernames = self.credentials.Username
        
    def pull(self):
        driver = self.driver
        wait = self.wait
        qtr = self.qtr
        yr = self.yr
        invoices_obtained=[]
        password = grabber.credentials.iloc[0,1]
        mapper = pd.read_excel(Grabloid.driver_path+'automation_parameters.xlsx',usecols='D,E',sheet_name='Alabama')
        mapper = dict(zip(mapper.iloc[:,0],mapper.iloc[:,1]))
        driver.implicitly_wait(15)
        for account in grabber.usernames:
            overall_success = 0
            while overall_success==0:
                try:
                    #initialize empty ndc list 
                    ndcs = []
                    driver.get('https://www.medicaid.alabamaservices.org/ALPortal/')
                    #Move to the drop down, hover and click "Secure Site"
                    drop_down =wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@title="Account"]')))
                    secure_site = driver.find_element_by_xpath('//a[@title="Secure Site"]')
                    ActionChains(driver).move_to_element(drop_down).move_to_element(secure_site).click().perform()
                    user = driver.find_element_by_xpath('//input[contains(@name,"userName")]')
                    user.send_keys(account)
                    pw = driver.find_element_by_xpath('//input[contains(@name,"password")]')
                    pw.send_keys(password)
                    login_button = driver.find_element_by_xpath('//a[contains(text(),"login")]')
                    login_button.click()
                    
                    #Move to trade files, hover and click for invoics
                    trade_files = wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@title="Trade Files"]')))
                    invoices = driver.find_element_by_xpath('//a[@title="Download"]')    
                    ActionChains(driver).move_to_element(trade_files).move_to_element(invoices).click().perform()    
                    print('A')
                    #Drop down menu permits selection of invoice
                    #type.  Get options and iterate through. 
                    
                    types = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//select[contains(@name,"TransactionType")]')))    
                    types_select = lambda: Select(types())   
                    types_to_get = [1,2,3]
                    print('B')
                    for report in types_to_get:
                        print('B1')
                        types_select().select_by_index(report)  
                        print('B2')
                        search_button = lambda: driver.find_element_by_xpath('//a[@title="Search using the specified criteria"]')  
                        print('B3')
                        canary = search_button()
                        search_button().click()
                        print('B4')
                        try:
                            print('B5')
                            alert = driver.switch_to.alert
                            alert.accept()
                        except NoAlertPresentException as ex:
                            print(ex.stacktrace)
                            print(ex.msg)
                            pass
                        wait.until(EC.staleness_of(canary))
                        print('C')
                        if report !=1:
                            invoice_period = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[contains(@name,"InvoicePeriod")]')))
                            invoice_period.clear()            
                            invoice_period.send_keys(str(qtr)+'/'+str(yr))
                            search_button().click()
                            files = driver.find_elements_by_xpath('//tr[contains(@class,"iC_DataList")]//td[2]')[2:]     
                            file_counter = 0
                            print('D')
                            for file in files:
                                file_counter +=1
                                file.click()
                                while len(os.listdir())==0:
                                    time.sleep(1)
                                while any(map((lambda x: 'RBT' in x), os.listdir()))==False:
                                    time.sleep(1)
                                while any(map((lambda x: 'crdownload' in x),os.listdir())) or any(map((lambda x: 'tmp' in x),os.listdir())):
                                    time.sleep(1)
                                file = os.listdir()[0]
                                label_code = file.split('.')[1]            
                                print('E')
                                if file[-3:]=='pdf':
                                    name = label_code+'_'+str(yr)+'Q'+str(qtr)+file[-4:]
                                else:
                                    with open(file) as ax:
                                        lines = ax.readlines()
                                    for ndc in list(set([x[6:17] for x in lines])):
                                        ndcs.append(ndc)
                                    name = label_code+'_'+str(yr)+'Q'+str(qtr)+'.txt'
                                    print(f'NDCS to get are',ndcs)
                                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Alabama\\CMS\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                                if os.path.exists(path)==False:
                                    os.makedirs(path)
                                else:
                                    pass
                                print('F')
                                if name in os.listdir(path):
                                    name = name.replace('.',f'_{file_counter}.')
                                invoices_obtained.append(name)
                                shutil.move(file,path+name)
                                time.sleep(8)
                                try:
                                    alert = driver.switch_to.alert
                                    alert.accept()
                                except NoAlertPresentException as ex:
                                    print(ex.stacktrace)
                                    print(ex.msg)
                                    pass
                                
                                
                        else:
                            file = driver.find_element_by_xpath('//tr[@class="iC_DataListItem"]//td[2]')            
                            file.click()
                            time.sleep(3)
                            while any(map((lambda x: 'RBT' in x), os.listdir()))==False:
                                time.sleep(1)
                            while any(map((lambda x: 'crdownload' in x),os.listdir())):
                                time.sleep(1)
                            for file in os.listdir():
                                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Alabama\\CMS\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                                if os.path.exists(path)==False:
                                    os.makedirs(path)
                                name = 'Invoice Cover Letter.pdf'
                                shutil.move(file,path+name)
                                invoices_obtained.append(name)
                        #if the file is the rtf format use it get get NDCS
                        if file[-4:]=='.rtf':
                            #now get the cld for each NDC
                            switch = 0
                            trade_files = wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@title="Trade Files"]')))
                            cld = driver.find_element_by_xpath('//a[@title="Claim Level Detail"]')
                            while switch==0:
                                ActionChains(driver).move_to_element(trade_files).move_to_element(cld).click().perform() 
                                time.sleep(3)
                                counter = 0 
                                try:
                                    alert = driver.switch_to.alert
                                    alert.accept()
                                    counter +=1
                                    time.sleep(1*counter)
                                except NoAlertPresentException as ex:
                                    switch=1
                            ndc_box = lambda: driver.find_element_by_xpath('//input[contains(@id,"CriteriaPanel_NDC")]')
                            invoice_period = lambda: driver.find_element_by_xpath('//input[contains(@name,"InvoiceCycle")]')
                            type_drop_down = lambda: driver.find_element_by_xpath('//select[contains(@name,"InvoiceType")]')                    
                            type_drop_down_select = lambda: Select(type_drop_down())
                            options = [x.text.strip() for x in type_drop_down_select().options]     
                            wait2 = WebDriverWait(driver,5)
                            ndcs = list(set(ndcs))
                            for option in options:
                                print(f'Selecting {option}')
                                type_drop_down_select().select_by_visible_text(option)
                                master_frame = pd.DataFrame()
                                for i, drug in enumerate(ndcs):
                                    print(f'Working on {drug}')
                                    cont_flag = 0
                                    ndc_box().clear()
                                    ndc_box().send_keys(drug)
                                    invoice_period().send_keys(str(qtr)+'/'+str(yr))
                                    print('Date entered')
                                    search_button = driver.find_element_by_xpath('//a[@title="Search using the specified criteria"]')
                                    print('Searching for records')
                                    search_button.click()
                                    print('Button clicked')
                                    wait.until(EC.staleness_of((search_button)))
                                    download_link = wait.until(EC.element_to_be_clickable((By.XPATH,'//a[text()="Download File"]')))
                                    try:
                                        success_flag = 0 
                                        try:
                                            driver.switch_to.default_content
                                        except:
                                            pass
                                        download_link.click()
                                        try:
                                            if download_link.get_attribute("disabled")=="true":
                                                print('Download link disabled, no data')
                                                continue
                                        except:
                                            print('Download link button not disabled')
                                        counter = 0
                                        print('Checking file is not crd or tmp')        
                                        while any((map((lambda x: 'crdownload' in x), os.listdir()))) or any(map((lambda x: 'tmp' in x),os.listdir())):
                                            time.sleep(1)
                                        while 'ClaimLevelDetail.csv' not in os.listdir():
                                            time.sleep(1)           
                                        flag = 0
                                        while flag == 0:
                                            try:
                                                a = open('ClaimLevelDetail.csv')
                                                flag=1
                                                a.close()
                                            except PermissionError as ex:
                                                flag = 0
                                                pass
                                        if i == 0:
                                            skip = 6
                                        else:
                                            skip = 7
                                        temp = pd.read_csv('ClaimLevelDetail.csv',usecols=list(range(16)),skiprows=6,engine='python', dtype=str)
                                        for col in temp.columns:
                                            if any(map((lambda x: '=' in x),temp[col]))==True:
                                                temp[col] = temp[col].str.replace('=','').str.replace('"','')
                                        meta_data = pd.read_csv('ClaimLevelDetail.csv',usecols=[0,1],nrows=5,header=None,names=['Field','Value'],engine='python')                  
                                        temp['NDC'] = ''.join(meta_data.Value[1].split('-'))
                                        temp['Program'] = meta_data['Field'][0]
                                        dp = len(temp)-1
                                        temp = temp.drop([dp])
                                        print('Temp dataframe created')
                                        master_frame = master_frame.append(temp)
                                        print('Temp dataframe appended to master')
                                        os.remove('ClaimLevelDetail.csv')
                                        print('Generic file removed')
                                        print(f"{'-'*15}\n")
                                    except:
                                        print('Some error')
                                flex_code = mapper[option]
                                file_name = f'AL_{flex_code}_{qtr}Q{yr}_{label_code}.csv'
                                print(f'File is named {file_name}')
                                master_frame.to_csv(file_name,index=False)
                                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Alabama\\'+option+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                                if os.path.exists(path)==False:
                                    os.makedirs(path)
                                else:
                                    pass
                                shutil.move(file_name,path+file_name)
                                print('File moved to folder, moving on\n')
                                print("----------------------------")
                    overall_success = 1
                except Exception as ex:
                    print(ex)    
                    print('An error occured')
                    pass
                        
                    
            drop_down =wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@title="Account"]')))
            log_off = driver.find_element_by_xpath('//a[@title="Logoff"]')
            log_off_able = 0
            while log_off_able==0:
                ActionChains(driver).move_to_element(drop_down).move_to_element(log_off).click().perform() 
                time.sleep(3)
                try:
                    alert = driver.switch_to.alert
                    alert.accept()
                except NoAlertPresentException as ex:
                    log_off_able=1
                    pass
        driver.close()
        os.chdir('O:\\')

        return invoices_obtained

    def morph_cld(self):
        path = 'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Alabama\\'
        for root, folders, files in os.walk(path):
            for file in files:
                try:
                    if file[-3:]=='csv':
                        df_path = root+'\\'+file
                        df = pd.read_csv(df_path,dtype=str)
                        new_name = root+'\\'+file[:-4]+'.xlsx'
                        df.to_excel(new_name, index=False)
                        os.remove(root+'\\'+file)
                except EmptyDataError as err:
                    data = ['This file was empty']
                    empty_df = pd.DataFrame(data)
                    new_name = root+'\\'+file[:-3]+'xlsx'
                    empty_df.to_excel(new_name, engine='xlsxwriter')
@push_note(__file__)                     
def main():
    grabber = Alabama_Grabloid()
    invoices = grabber.pull()
    grabber.morph_cld()
    grabber.send_message(invoices)
if __name__=='__main__':
    main()
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    