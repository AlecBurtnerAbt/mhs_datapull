# -*- coding: utf-8 -*-
"""
Created on Fri Aug 24 09:06:01 2018

@author: c252059
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
import multiprocessing as mp
from multiprocessing.pool import Pool
from grabloid import Grabloid, push_note
from pushover_complete import PushoverAPI


class IllinoisGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Illinois')

    def pull(self):
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Illinois', usecols='D,E',dtype='str')
        mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
        time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
        yr = time_stuff.iloc[0,0]
        qtr = time_stuff.iloc[0,1]
        yq=str(yr)+str(qtr)
        login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Illinois', usecols=[0,1],dtype='str')
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        driver.get(r'https://rsp.ilgov.emdeon.com/RebateServicesPortal/login/home?goto=http://rsp.ilgov.emdeon.com/RebateServicesPortal/')
        wait2 = WebDriverWait(driver,2)
        driver.implicitly_wait(20)
        #find username and password and pass the login credentials
        
        user = driver.find_element_by_xpath('//input[@id="username"]')
        user.send_keys(username)
        pw = driver.find_element_by_xpath('//input[@id="password"]')
        pw.send_keys(password)
        login = driver.find_element_by_xpath('//input[@value="Login"]')
        login.click()
        
        #Now to navigate past the next page
        
        accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Accept"]')))
        accept.click()
        
        #Now to get to the invoices page
        invoices = driver.find_element_by_xpath('//a[text()="Invoices"]')
        invoices.click()
        
        #We have labeler, state, and type
        
        labeler = lambda: driver.find_element_by_xpath('//select[@id="labeler"]')
        labeler_select = lambda: Select(labeler())
        options = [x.text for x in labeler_select().options]
        options = options[1:]
        
        types = lambda: driver.find_element_by_xpath('//select[@name="docType"]')
        types_select = lambda: Select(types())
        
        time_stamp = lambda: driver.find_element_by_xpath('//input[@id="period"]')
        time_stamp().send_keys(yq)
        
        report_values = [x.get_attribute('value') for x in driver.find_elements_by_xpath('//select[@id="docType"]/option')]
        reports = [x.text for x in driver.find_elements_by_xpath('//select[@id="docType"]/option')]
        mapper2 = dict(zip(report_values,reports))
        
        #now begin looping
        invoices_obtained = []
        master_dict = dict.fromkeys(options)
        for label in options:
            labeler_select().select_by_visible_text(label)
            for report in list(mapper.keys()):
                report_dict = {}
                submit_button = driver.find_element_by_xpath('//input[@value="Submit"]')
                types_select().select_by_visible_text(report)
                time_stamp().send_keys(yq)
                submit_button.click()
                report_value = driver.find_element_by_xpath('//select[@id="docType"]/option[text()="{}"]'.format(report)).get_attribute('value')
                wait.until(EC.staleness_of(submit_button))
                try:
                    downloads = driver.find_elements_by_xpath('//table[@id="invoiceResults"]//a[@title="Download"]')
                except TimeoutException as err:
                    invoices_obtained.append([label,report,"NOT OBTAINED! NOT AVAILABLE YET!"])
                if len(downloads)==0:
                    invoices_obtained.append([label,report,"NOT OBTAINED! NOT AVAILABLE YET!"])
                    continue
                else:
                    pass
                for i,link in enumerate(downloads):
                    state = 'IL'
                    #CMS format always listed first so 0=.txt and 1=.pdf
                    if i==0:
                        file_type = '.txt'
                    else:
                        file_type = '.pdf'
                    #click the link to download the file
                    link.click()
                    print(f'Link clicked for {file_type}')
                    #an alert pops up, switch to it and accept it
                    try:
                        alert = driver.switch_to.alert
                        alert.accept()
                    except NoAlertPresentException:
                        pass
                    #depending on the file type the name will be different, below builds the different file name formats
                    if i==0:
                        file_name = '-'.join([state,label,yq,report_value])+file_type
                    else:
                        file_name = ''.join([label,report_value,yq])+file_type
                    #rename for the pre-processor tool
                    new_name = 'IL_{}_{}Q{}_{}{}'.format(mapper2[report_value],qtr,yr,label,file_type)
                    #wait until file is obtained
                    print('Waiting for file to arrive in temp folder')
                    while file_name in os.listdir()==False:
                        time.sleep(1)
                    #now read the file if it is a .txt
                    print('File downloaded')
                    if file_type =='.txt':
                        print('Attempting to read file')
                        read_success = 0
                        while read_success == 0:
                            try:
                                with open(file_name) as F:
                                    lines = F.readlines()
                                    ndcs = list(set([x[6:17] for x in lines]))
                                    report_dict.update({report:ndcs})
                                    read_success=1
                            except:
                                pass
                        print('File read')
                    #build the path for the file, 
                    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Illinois\\{}\\{}\\{}\\'.format(yr,qtr,mapper2[report_value])                
                    if os.path.exists(path)==False:
                        os.makedirs(path)
                    else:
                        pass
                    #was having issues with Python not finding the file despite it being in the folder,
                    #so have the transfer flag while loop to ensure that the transfer happens
                    transfer_flag = 0
                    while transfer_flag ==0:
                        try:
                            shutil.move(file_name,path+new_name)
                            invoices_obtained.append(' '.join([label,report,file_type]))
                            transfer_flag=1
                        except NameError as err:
                            time.sleep(1)
                        except FileNotFoundError as err:
                            time.sleep(1)
                        except PermissionError as err:
                            time.sleep(1)
            master_dict.update({label:report_dict})         
        driver.close()
        return yq, username, password, master_dict, invoices_obtained
    
    
    
    
    
    def download_reports(self):
        os.chdir('C:/Users/')
        chromeOptions = webdriver.ChromeOptions()
        prefs = {'download.default_directory':f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder\\{self.script}\\',
                 'plugins.always_open_pdf_externally':True,
                 'download.prompt_for_download':False}
        chromeOptions.add_experimental_option('prefs',prefs)
        driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\chromedriver.exe')
        os.chdir(self.temp_folder_path)
        yr = self.yr
        qtr = self.qtr
        login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name=f'{self.script}', usecols='A,B',dtype='str')
        username = self.credentials.iloc[0,0]
        password = self.credentials.iloc[0,1]
        yq=str(yr)+str(qtr)
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name=f'{self.script}', usecols='D,E',dtype='str')
        mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
        
        
        
        
        #Login with provided credentials
        driver.get(r'https://rsp.ilgov.emdeon.com/RebateServicesPortal/login/home?goto=http://rsp.ilgov.emdeon.com/RebateServicesPortal/')
        
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
        pages = driver.find_elements_by_xpath('//div[@class="dataTables_paginate pagination"]/a[@class="step"]')
        num_pages = len(pages)+1
        #There is a long "flash to bang" for the preparation of the requested reports, 
        #so this will have to be a two part operation. 
        master_df = pd.DataFrame()
        if len(rows)==0:
            pass
        else:
            for page in range(num_pages):
                print(f'Working on page {page+1}')
                #Redefine rows for each page
                rows = driver.find_elements_by_xpath('//table[@id="reportsResults"]/tbody/tr')
                rows = [row for row in rows if checker(row,'td//a//span[text()="Download Report"]')==True]
                #get names of files and their links
                names = [x.find_element_by_xpath('td[1]').text for x in rows]
                links = [x.find_element_by_xpath('td//a[@href="#"]') for x in rows]
                print('Looping through names and links...')              
                for name, link in zip(names, links):
                    #get info for file name
                    ndc = name.split(' ')[7]
                    state = name.split(' ')[8]
                    program = name.split(' ')[10]
                    value = mapper[program]
                    first_half = '_'.join(name.split(' ')[:5])
                    second_half = '-'.join(name.split(' ')[-4:]).replace(program,mapper[program])
                    download_name = '-'.join([first_half,second_half])+'.xls'
                    #download the file
                    flag = 0
                    print(f'Downloading {download_name}')
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
                    print('Download success')
                    read_flag =0                   
                    while read_flag==0:
                        print('Reading file...')
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
                    print(f'{download_name} read and appended to master df!')
                if page != num_pages-1:
                    print(f'Getting page {page}')
                    next_page = driver.find_element_by_xpath('//div[@class="dataTables_paginate pagination"]/a[@class="nextLink"]')
                    next_page.click()
                    wait.until(EC.staleness_of(next_page))
                else:
                    print('Done!')

            frames = []
            splitters = master_df.Program.unique().tolist()  
            for splitter in splitters:
                frame = master_df[master_df['Program']==splitter]
                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Illinois\\'+splitter+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                file_name = 'IL_'+splitter+'_'+str(qtr)+'Q'+str(yr)+'.xlsx'
                if os.path.exists(path)==False:
                    os.makedirs(path)
                else:
                    pass
                os.chdir(path)
                frame.to_excel(file_name, engine='xlsxwriter',index=False)
            #now delete all the files that have been downloaded
            deletes = lambda: driver.find_elements_by_xpath('//table[@id="reportsResults"]//a[@title="Delete"][@class="btn"]')
            while len(deletes())>0:
                for i in range(len(deletes())):
                    canary = driver.find_element_by_xpath('//input[@id="reportSub"]')
                    deletes()[0].click()
                    alert = driver.switch_to.alert
                    alert.accept()
                    wait.until(EC.staleness_of(canary))
                    
            
        driver.close()
        os.chdir('O:\\')
        os.removedirs(self.temp_path_folder)   
        
    def make_chunks(self,master_dict):
        #Break the information for each report down into 
        reports = []
        for key in master_dict.keys():
            for key2 in master_dict[key].keys():
                for value in master_dict[key][key2]:
                    if len(value)==0:
                        continue
                    else:
                        report = (key,key2,value)
                        reports.append(report)
        n = round(len(reports)/(mp.cpu_count()-1))
        chunks = [reports[x:x+n] for x in range(0,len(reports),n)]
        return chunks

    
    
def getReports(num,chunk):
    print(f'Hello from {num}')
    print('Working on chunk: '+str(num))
    time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
    yr = time_stuff.iloc[0,0]
    qtr = time_stuff.iloc[0,1]
    yq=str(yr)+str(qtr)
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Illinois', usecols='A,B',dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    os.chdir('C:/Users/')
    chromeOptions = webdriver.ChromeOptions()
    prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder\\Illinois\\',
         'plugins.always_open_pdf_externally':True,
         'download.prompt_for_download':False}
    chromeOptions.add_experimental_option('prefs',prefs)
    driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\chromedriver.exe')
    os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder\\Illinois\\')
    #Login with provided credentials
    driver.get(r'https://rsp.ilgov.emdeon.com/RebateServicesPortal/login/home?goto=http://rsp.ilgov.emdeon.com/RebateServicesPortal/')
    
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
    reports_tab = wait.until(EC.element_to_be_clickable((By.XPATH,'//a[text()="Reports"]')))       
    reports_tab.click()           
    #Now starting iterating through the chunk
    for label, program, ndc in chunk:
        success = 0
        report = wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="reportList"]')))
        report_select = Select(report)
        while success==0:
            try:
                report = driver.find_element_by_xpath('//select[@name="stateReportId"]')
                select_report = Select(report)     
                if 'jcode' in program.lower():
                    select_report.select_by_index(3)
                else:
                    select_report.select_by_index(2)
                
                ndc_in = driver.find_element_by_xpath('//input[@name="ndc"]')
                ndc_in.send_keys(ndc)
                
                docType = driver.find_element_by_xpath('//select[@name="docType"]')
                select_docType = Select(docType)
                select_docType.select_by_visible_text(program.replace('_',' '))
                
                rpu = driver.find_element_by_xpath('//input[@name="rpuStart"]')
                rpu.send_keys(yq)
                
                submit_button= driver.find_element_by_xpath('//input[@value="Submit"]')
                submit_button.click()
                
                accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Accept"]')))
                accept.click()
            except:
                driver.refresh()
                
            wait.until(EC.staleness_of(accept))
            success=1
    driver.close()    



def multi_grabber(i, chunks):             
    processes = [mp.Process(target=getReports,args=(i,chunk)) for i,chunk in enumerate(chunks)]
    for p in processes:
        p.start()       
    for p in processes:
        p.join() 
        
@push_note(__file__)        
def main():
    grabber = IllinoisGrabloid()
    yq, username, password, master_dict, invoices = grabber.pull()      
    chunks = grabber.make_chunks(master_dict)
    multi_grabber(enumerate(chunks))
  



if __name__=='__main__':   
    main()








