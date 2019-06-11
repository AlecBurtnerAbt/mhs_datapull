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
import multiprocessing as mp
from grabloid import Grabloid


class PennsylvaniaGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Pennsylvania')


    def stepOne(self):

        yr = self.yr
        qtr = self.qtr
        login_credentials = self.credentials
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        yq=str(yr)+str(qtr)
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Pennsylvania', usecols='D,E',dtype='str')
        mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
        driver = self.driver
        #Login with provided credentials
        driver.get('https://rsp.pagov.changehealthcare.com/RebateServicesPortal/login/home?goto=http://rsp.pagov.changehealthcare.com/RebateServicesPortal/')   
        
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
        invoices_obtained=[]
        invoices = driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div/ul/li[2]/a')
        invoices.click()
        code_dropdown = lambda: driver.find_element_by_id('labeler')
        code_select = lambda: Select(code_dropdown())
        codes = [x.text for x in code_select().options][1:]
        type_dropdown = lambda: WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.ID,'docType')))
        type_select = lambda: Select(type_dropdown())
        master_dict = dict.fromkeys(codes)
        time_stamp = lambda: driver.find_element_by_xpath('//input[@name="period"]')
        
        
        
        for code in codes:
            code_select().select_by_visible_text(code)
            print('Selecting '+str(code))
            report_dict = {}
            for report in list(mapper.keys()):
                if report == "DME" or report == "SR":
                    continue
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
                links = lambda: driver.find_elements_by_xpath('//a[@title="Download"]')
                for i,link in enumerate(links()):
                    success_flag =0
                    if i ==0:
                        file_type = '.txt'
                    else:
                        file_type = '.pdf'
                    file = 'PA-'+code+'-'+yq+'-'+report+file_type
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
                        read_flag = 0
                        while read_flag==0:
                            try:
                                with open(file) as ax:
                                    lines = ax.readlines()
                                    ndcs = list(set([x[6:17] for x in lines]))
                                    report_dict.update({report:ndcs})
                                    read_flag=1
                            except PermissionError as err:
                                time.sleep(1)
                    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Pennsylvania\\'+report+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                    if os.path.exists(path)==False:
                        os.makedirs(path)
                    else:
                        pass
                    file_name = 'PA_{}_{}Q{}_{}{}'.format(report,qtr,yr,code,file_type)
                    invoices_obtained.append(file_name)
                    shutil.move(file,path+file_name)
            master_dict.update({code:report_dict})
        driver.close()
        return yq, username, password, master_dict, invoices_obtained
    
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
        import math
        n = math.ceil(len(reports)/6)
        chunks = [reports[x:x+n] for x in range(0,len(reports),n)]
        return chunks
    #write the function for multi threading
    
    
    
    
        
    def download_reports(self):
        os.chdir('C:/Users/')
        chromeOptions = webdriver.ChromeOptions()
        prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder',
                 'plugins.always_open_pdf_externally':True,
                 'download.prompt_for_download':False}
        chromeOptions.add_experimental_option('prefs',prefs)
        driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'C:\chromedriver.exe')
        os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
        #for file in os.listdir():
            #os.remove(file)
        time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
        yr = time_stuff.iloc[0,0]
        qtr = time_stuff.iloc[0,1]
        login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Pennsylvania', usecols='A,B',dtype='str')
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        yq=str(yr)+str(qtr)
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Pennsylvania', usecols='D,E',dtype='str')
        mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
        
        
        
        
        #Login with provided credentials
        driver.get('https://rsp.pagov.changehealthcare.com/RebateServicesPortal/login/home?goto=http://rsp.pagov.changehealthcare.com/RebateServicesPortal/')   
        
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
            program = name.split(' ')[10]
            value = mapper[program]
            first_half = '_'.join(name.split(' ')[:5])
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
        splitters = master_df.Program.unique().tolist()  
        for splitter in splitters:
            frame = master_df[master_df['Program']==splitter]
            path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Pennsylvania\\'+splitter+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
            try:
            	program_code_name = mapper[splitter]
            except:
            	program_code_name = splitter
            file_name = 'PA_'+program_code_name+'_'+str(qtr)+'Q'+str(yr)+'.xlsx'
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
        os.chdir('O:\\')



def getReports(num,chunk):
        print('Working on chunk: '+str(num))
        time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
        yr = time_stuff.iloc[0,0]
        qtr = time_stuff.iloc[0,1]
        yq=str(yr)+str(qtr)
        login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Pennsylvania', usecols='A,B',dtype='str')
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        os.chdir('C:/Users/')
        chromeOptions = webdriver.ChromeOptions()
        prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder',
             'plugins.always_open_pdf_externally':True,
             'download.prompt_for_download':False}
        chromeOptions.add_experimental_option('prefs',prefs)
        #chromeOptions.add_argument('--headless')
        #chromeOptions.add_argument('--disable-gpu')
        driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'C:\chromedriver.exe')
        os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder')
        #Login with provided credentials
        driver.get('https://rsp.pagov.changehealthcare.com/RebateServicesPortal/login/home?goto=http://rsp.pagov.changehealthcare.com/RebateServicesPortal/')   
        
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
        report = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="reportList"]')))
        report_select = lambda: Select(report())
        #Now starting iterating through the chunk
        for label, program, ndc in chunk:
            success = 0
            program = program.strip()
            while success==0:
                try:
                    report = wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@name="stateReportId"]')))
                    select_report = Select(report)        
                    select_report.select_by_index(2)
                    
                    ndc_in = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@name="ndc"]')))
                    ndc_in.send_keys(ndc)
                    
                    docType = wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@name="docType"]')))
                    select_docType = Select(docType)
                    select_docType.select_by_visible_text(program.replace('_',' '))
                    
                    rpu = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@name="rpuStart"]')))
                    rpu.send_keys(yq)
                    
                    submit_button= wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Submit"]')))
                    submit_button.click()
                    
                    accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Accept"]')))
                    accept.click()
                except NoSuchElementException as ex:
                    print(ex.msg)
                    print(ex.stacktrace)
                    driver.refresh()
                except TimeoutException as ex:
                    print(ex.msg)
                    print(ex.stacktrace)
                    driver.refresh()
                    
                wait.until(EC.staleness_of(accept))
                soup = BeautifulSoup(driver.page_source,'html.parser')
                Reports = [x.text.strip() for x in soup.find_all('td')]
                if any(map((lambda x: (ndc+' PA '+yq+' '+program) in x),Reports)):
                    success=1
                else:
                    pass
        driver.close()
        
def main():
    grabber = PennsylvaniaGrabloid()
    yq, username, password, master_dict, invoices = grabber.stepOne()      
    chunks = grabber.make_chunks(master_dict)
    processes = [mp.Process(target=getReports,args=(i,chunk)) for i,chunk in enumerate(chunks)]
    for p in processes:
        p.start()       
    for p in processes:
        p.join()  
    grabber.download_reports()


       
if __name__=='__main__':
    main()
   




