# -*- coding: utf-8 -*-
"""
Created on Thu Aug 30 09:28:25 2018

@author: C252059
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Aug 30 08:52:03 2018

@author: C252059
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Aug 28 09:09:59 2018

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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoAlertPresentException, InvalidElementStateException
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
from xlrd.biffh import XLRDError
import xlsxwriter as xl
import requests
from requests.auth import HTTPBasicAuth
from grabloid import Grabloid


class WyomingGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Wyoming')
        self.usernames = self.credentials.Username
        
        
    def pull(self):
        driver = self.driver
        wait = self.wait
        time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
        yr = time_stuff.iloc[0,0]
        qtr = time_stuff.iloc[0,1]
        login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Wyoming', usecols='A,B',dtype='str')
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        yq=str(yr)+str(qtr)
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Wyoming', usecols='D,E',dtype='str')
        mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
        
        #Login with provided credentials
        driver.get('https://rsp.wygov.changehealthcare.com/RebateServicesPortal/login/home?goto=http://rsp.wygov.changehealthcare.com/RebateServicesPortal')   
        
        #Now login
        
        user = driver.find_element_by_xpath('//input[@id="username"]')
        user.send_keys(username)
        pass_word = driver.find_element_by_id('password')
        pass_word.send_keys(password)
        login = driver.find_element_by_id('submit')
        login.click()
        
        wait = WebDriverWait(driver,10)
        wait2 = WebDriverWait(driver,2)
        accept = wait.until(EC.element_to_be_clickable((By.ID,'terms')))
        accept.click()
        
        #invoice stuff is below this
        
        invoices = driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div/ul/li[2]/a')
        invoices.click()
        code_dropdown = lambda: driver.find_element_by_id('labeler')
        code_select = lambda: Select(code_dropdown())
        codes = [x.text for x in code_select().options][1:]
        type_dropdown = lambda: WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.ID,'docType')))
        type_select = lambda: Select(type_dropdown())
        master_list = []
        time_stamp = lambda: driver.find_element_by_xpath('//input[@name="period"]')
        values = driver.find_elements_by_xpath('//select[@id="docType"]//option')
        values = [x.get_attribute('value') for x in values if int(x.get_attribute('value'))>1]
        
        for code in codes:
            code_select().select_by_visible_text(code)
            print('Selecting '+str(code))
            ndcs = []
            for report,value in zip(list(mapper.keys()),values):
                
                type_select().select_by_visible_text(report)
                time.sleep(1)
                if ' ' in report:
                    report = report.replace(' ','_')
                else:
                    pass
                print('Selecting '+report)
                try:
                    time_stamp().clear()
                    time_stamp().send_keys(yq)
                except InvalidElementStateException as ex:
                    pass
                submit_button = driver.find_element_by_xpath('//input[@type="submit"]')
                print('Requesting file.')
                submit_button.click()
                wait.until(EC.staleness_of(submit_button))
                success=0
                while success ==0:
                    try:
                        error = wait2.until(EC.presence_of_element_located((By.XPATH,'//li[contains(text(),"An error")]')))
                        print('Website error! Moving back')
                        driver.back()
                        type_select().select_by_visible_text(report)
                        submit_button = driver.find_element_by_xpath('//input[@type="submit"]')
                        print('Requesting file.')
                        submit_button.click()
                        
                    except TimeoutException as ex:
                        success=1
                print('Files returned.')
                wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@title="Download"]')))
                links = lambda: driver.find_elements_by_xpath('//a[@title="Download"]')
                for i,link in enumerate(links()):
                    success_flag =0
                    if i ==0:
                        file_type = '.txt'
                    else:
                        file_type = '.pdf'
                    file = 'WY-'+code+'-'+yq+'-'+value+file_type
                    print('Downloading file '+str(i+1))
                    reset_counter=0
                    while success_flag ==0:              
                        links()[i].click()
                        try:
                            alert = driver.switch_to.alert
                            alert.accept()
                        except NoAlertPresentException as ex:
                            pass
                        try:
                            error = wait2.until(EC.presence_of_element_located((By.XPATH,'//li[contains(text(),"An error")]')))
                            print('Website error! Moving back')
                            driver.back()
                            reset_counter+=1
                            time.sleep(reset_counter*1.5)
                            reset_flag = 1
                        except TimeoutException as ex:
                            reset_flag=0
                            pass
                        counter = 0
                        if reset_flag ==1:
                            pass
                        else:
                            while file not in os.listdir() and counter<10:
                                time.sleep(1)
                                counter+=1
                            if file in os.listdir():
                                success_flag=1
                            else:
                                pass
                    if file_type == '.txt':
                        read_success=0
                        while read_success==0:
                            try:
                                with open(file) as ax:
                    
                                    lines = ax.readlines()
                                    xxx = list(set([x[6:17] for x in lines]))
                                    [master_list.append((report,code,x)) for x in xxx]
                                    read_success=1
                            except PermissionError as ex:
                                pass
                                
                    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Wyoming\\'+report+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                    if os.path.exists(path)==False:
                        os.makedirs(path)
                    else:
                        pass
                    file_name = 'WY_'+report+'_'+code+'_'+str(qtr)+'Q'+str(yr)+file_type
                    shutil.move(file,path+file_name)
        
        
             #############################################CLD below
        
        reports_tab = driver.find_element_by_xpath('//a[text()="Reports"]')       
        reports_tab.click()                
        
        not_ready = []
        master_df = pd.DataFrame()
        
        #master list is composed of tuples which contain (program, label code, NDC)
        
        for item in master_list:
            program = item[0]
            label_code = item[1]
            NDC = item[2]
            reports = wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="reportList"]')))
            reports_select = Select(reports)
            #make sure no bad data is passed which will crash the program
            if len(NDC) <11:
                continue
            if program=='JCode':
                reports_select.select_by_index(2)
            else:
                reports_select.select_by_index(1)
            claim_counter=0
            claim_select =0
            while claim_counter <10 and claim_select==0:
                try:
                    claim_type_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="docType"]')))
                    claim_type_dropdown_select = Select(claim_type_dropdown)
                    claim_type_dropdown_select.select_by_visible_text(program)
                    claim_select=1
                except NoSuchElementException as ex:
                    counter+=1
                    time.sleep(1)
            ndc_in = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@id="NDCReport"]')))
            ndc_in.send_keys(NDC)
            date = driver.find_element_by_xpath('//input[@id="RpuStartReport"]')
            date.send_keys(yq)
            submit = driver.find_element_by_xpath('//input[@id="reportSub"]')
            submit.click()
            warning_button = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@id="acceptWarning"]')))
            warning_button.click()
    

    def download_reports(self):
        driver = self.driver
        #Login with provided credentials
        driver.get('https://rsp.wygov.changehealthcare.com/RebateServicesPortal/login/home?goto=http://rsp.wygov.changehealthcare.com/RebateServicesPortal')   
        
        #Now login
        
        user = driver.find_element_by_xpath('//input[@id="username"]')
        user.send_keys(username)
        pass_word = driver.find_element_by_id('password')
        pass_word.send_keys(password)
        login = driver.find_element_by_id('submit')
        login.click()
        
        wait = WebDriverWait(driver,10)
        wait2 = WebDriverWait(driver,2)
        accept = wait.until(EC.element_to_be_clickable((By.ID,'terms')))
        accept.click()
        reports_tab = driver.find_element_by_xpath('//a[text()="Reports"]')       
        reports_tab.click()                
        report = lambda: driver.find_element_by_xpath('//select[@id="reportList"]')
        report_select = lambda: Select(report())
        
        types = driver.find_element_by_xpath('//select[@id="docType"]')
        types_select = Select(types)
        programs = [x.text.replace(' ','_') for x in types_select.options]
        values = [x.get_attribute('value') for x in driver.find_elements_by_xpath('//select[@id="docType"]/option')]
        mapper = dict(zip(programs,values))    
        #Helper function to return boolean if report is ready
        def checker(element,xpath):
            try:
                EC.presence_of_element_located(element.find_element_by_xpath(xpath))
                return True
            except NoSuchElementException as ex:
                return False
        #Below is where the script finds the reports, downloads, and moves them
        rows = driver.find_elements_by_xpath('//table[@id="reportsResults"]/tbody/tr')
        rows = [row for row in rows if checker(row,'td//a//span[text()="Download Report"]')==True]
        
        #now that we have rows only for where reports are ready we can move forward
        names = [x.find_element_by_xpath('td[1]').text for x in rows]
        links = [x.find_element_by_xpath('td//a[@href="#"]') for x in rows]
        master_df = pd.DataFrame()
        
        for name, link in zip(names, links):
            #get info for file name
            ndc = name.split(' ')[7]
            state = name.split(' ')[8]
            if 'JCode' in name:
                program = 'JCode'
            else:
                program = name.split(' ')[10]
            value = mapper[program]
            #build the file name to look for in the download folder, 
            #is a different size depending on what kind of file it is
            if program != 'JCode':
                first_half = '_'.join(name.split(' ')[:5])
                second_half = '-'.join(name.split(' ')[-4:]).replace(program,mapper[program])
                download_name = '-'.join([first_half,second_half])+'.xls'
            else:
                first_half = '_'.join(name.split(' ')[:3])
                second_half = '-'.join(name.split(' ')[-4:]).replace(program,mapper[program])
                download_name = '-'.join([first_half,second_half])+'.xls'
            #download the file
            flag = 0
            while flag ==0:
                link.click()
                counter = 0
                while download_name not in os.listdir() and counter<21:
                    time.sleep(1)
                    counter+=1
                if download_name not in os.listdir():
                    pass
                else:
                    flag = 1
            read_flag =0
            while read_flag==0:
                try:
                    temp_df = pd.read_excel(download_name,skipfooter=3)
                    read_flag=1
                except PermissionError as err:
                    time.sleep(1)
            temp_df = temp_df.dropna(axis=0,how='all')
            if len(temp_df)==0:
                continue
            else:
                pass
            temp_df['NDC']= ndc
            temp_df['Program'] = program
            master_df = master_df.append(temp_df)
        frames = []
        splitters = master_df['Program'].unique().tolist()  
        for splitter in splitters:
            frame = master_df[master_df['Program']==splitter]
            path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Wyoming\\'+splitter+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
            file_name = 'WY_'+splitter+'_'+str(qtr)+'Q'+str(yr)+'.xlsx'
            if os.path.exists(path)==False:
                os.makedirs(path)
            else:
                pass
            os.chdir(path)
            frame.to_excel(file_name, engine='xlsxwriter',index=False)
        #now delete all the files that have been downloaded
        deletes = lambda: driver.find_elements_by_xpath('//table[@id="reportsResults"]//a[@title="Delete"][@class="btn"]')
        for i in range(len(deletes())):
            canary = driver.find_element_by_xpath('//input[@id="reportSub"]')
            deletes()[0].click()
            alert = driver.switch_to.alert
            alert.accept()
            wait.until(EC.staleness_of(canary))       
        driver.close()


if __name__=="__main__":
    grabber = WyomingGrabloid()
    grabber.download_reports()
    




