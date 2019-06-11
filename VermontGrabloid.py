# -*- coding: utf-8 -*-
"""
Created on Tue Dec  4 15:09:01 2018

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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
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
from grabloid import Grabloid, push_note




class Vermont_Grabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Vermont')
        
    def delete_old_reports(self):
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Vermont', usecols='D,E',dtype='str')
        mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
        time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
        yr = time_stuff.iloc[0,0]
        qtr = time_stuff.iloc[0,1]
        yq=str(yr)+str(qtr)
        login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Vermont', usecols=[0,1],dtype='str')
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
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
        reports_tab = driver.find_element_by_xpath('//a[text()="Reports"]')       
        reports_tab.click() 
        deletes = lambda: driver.find_elements_by_xpath('//table[@id="reportsResults"]//a[@title="Delete"][@class="btn"]')
        while len(deletes())>0:
            for i in range(len(deletes())):
                canary = driver.find_element_by_xpath('//input[@id="reportSub"]')
                deletes()[0].click()
                alert = driver.switch_to.alert
                alert.accept()
                wait.until(EC.staleness_of(canary))
                
    def get_invoices(self):
        driver = self.driver
        yr = self.yr
        qtr = self.qtr
        yq=str(yr)+str(qtr)
        username = self.credentials.iloc[0,0]
        password = self.credentials.iloc[0,1]
    
        driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
        user = driver.find_element_by_id('username')
        user.send_keys(username)
        pass_word = driver.find_element_by_id('password')
        pass_word.send_keys(password)
        login = driver.find_element_by_id('submit')
        login.click()
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Vermont', usecols='D,E',dtype='str')
        mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
        wait = WebDriverWait(driver,10)
        try:
            accept = wait.until(EC.element_to_be_clickable((By.ID,'terms')))
            accept.click()
        except TimeoutException as ex:
            print('Dont have to accept twice dude')
        
        #invoice stuff is below this
        
        invoices = driver.find_element_by_xpath('//a[text()="Invoices"]')
        invoices.click()
        code_dropdown = lambda: driver.find_element_by_id('labeler')
        code_select = lambda: Select(code_dropdown())
        codes = code_select().options
        type_dropdown = lambda: WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.ID,'docType')))
        type_select = lambda: Select(type_dropdown())
        types = type_select().options
        report_ndcs = []
        codes = [item.text for item in codes[1:]]
        types = [item.text for item in types[1:]]
        dme_index = types.index([x for x in types if 'DME' in x][0])
        types[dme_index] = 'DME'
        medicare_wrap_index = types.index([x for x in types if 'Wrap' in x][0])
        types[medicare_wrap_index] = 'Medicare_Wrap'
        types_2 = []
        invoices_obtained = []
        invoices_to_get = list(itertools.product(codes,types))  
        reference_list = []
        values = [x.get_attribute('value') for x in type_select().options][1:]
        mapper2 = dict(zip(types,values))
        #helper function to check dates in results table
        def check():
            invoice_period = driver.find_element_by_xpath('//table[@id="invoiceResults"]/tbody/tr/td[4]')
            if invoice_period.text == yq:
                available =1
            else:
                available = 0
            return available
        #Helper function to check for error page
        def check_for_error():
            try:
                error_message = driver.find_element_by_xpath('//li[contains(text(),"An error has occurred")]')
                return 1
            except:
                return 0
            
        for label, report in invoices_to_get:
            counter = 0
            available =0
            while available ==0:
                counter +=1
                code_select().select_by_visible_text(label)
                time.sleep(1)
                type_select().select_by_index(types.index(report)+1)
                time.sleep(1)
                submit = driver.find_element_by_xpath('//input[@id="invSubmit"]')
                print(f'Requesting report for {report} {label} \n')
                submit.click()
                wait.until(EC.staleness_of(submit))
                available = check()
                if available ==0:
                    driver.refresh()
                else:
                    pass
                time.sleep(counter*1.5)
            buttons = lambda: driver.find_elements_by_xpath('//a[@class="btn"][contains(@onclick,"download")]')
            for i, button in enumerate(buttons()):
                success = 0
                counter = 0
                continue_flag =0
                while success==0 and counter < 2:
                    try:
                        print('Clicking button')
                        buttons()[i].click()
                        print('Button clicked')
                        alert = driver.switch_to.alert
                        alert.accept()
                        print('Alert accepted')
                        if i == 0:
                            file_type='.txt'
                            print('File is a text file')
                        else:
                            file_type = '.pdf'
                            print('File is a pdf file')
                        if ' ' in report:
                            modified_report = report.replace(' ','_')
                            file_name = f'VT-{label}-{yq}-{modified_report}{file_type}'
                        else:
                            file_name = f'VT-{label}-{yq}-{report}{file_type}'
                        print(f'File name is {file_name}\n')
                        if check_for_error() == 1:
                            print('Error')
                            driver.refresh()
                            continue
                        else:
                            print('No Error Found')
                            pass        
                        counter = 0
                        while file_name not in os.listdir() and counter<30:
                            counter+=1
                            print(f'{30-counter} seconds remaining before failover operation')
                            time.sleep(1)
                        if counter>29:
                            print('File did not download within 30 seconds, retrying')
                            continue
                        else:
                            print('File downloaded')
                        #Now open the file and return the NDCs associated to the label code and program
                        if file_type =='.txt':
                            read_flag = 0
                            while read_flag==0:
                                try:
                                    with open(file_name) as f:
                                        lines = f.readlines()
                                        ndcs = list(set([line[6:17] for line in lines]))
                                        ndcs = [ndc for ndc in ndcs if len(ndc)>1]
                                        reference_list.append((label,report,ndcs))
                                    read_flag=1
                                except PermissionError as ex:
                                    pass
                        else:
                            pass
                        try:
                            report_value = mapper2[report]
                            flex_name = mapper[report_value]
                        except KeyError as err:
                            flex_name = report
                        print(f'Flex name is {flex_name}')
                        new_name = f'VT_{flex_name}_{qtr}Q{yr}_{label}{file_type}'
                        print(f'New name is {new_name}')
                        path =  f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Vermont\\{flex_name}\\{yr}\\Q{qtr}\\'
                        print(f'Moving to {path}')
                        if os.path.exists(path)==False:
                            os.makedirs(path)
                        shutil.move(file_name,path+new_name)
                        success=1
                    except:
                        counter +=1
                        print(f'Error occurred while obtaining invoice for {report} for labeler {label}')
                        driver.back()
                        code_select().select_by_visible_text(label)
                        time.sleep(1)
                        type_select().select_by_index(types.index(report)+1)
                        time.sleep(1)
                        submit = driver.find_element_by_xpath('//input[@id="invSubmit"]')
                        print(f'Requesting report for {report} {label} \n')
                        submit.click()
                        wait.until(EC.staleness_of(submit))
                    if counter >0:
                        continue_flag = 1
                        break
                if continue_flag ==1:
                    break
                else:
                    pass
            if continue_flag ==1:
                continue
            else:
                pass
        from collections import defaultdict
        master_dict = defaultdict(dict)        
        for label, report, ndcs in reference_list:
            if len(ndcs)>0:
                master_dict[label][report]=ndcs        
        driver.stop_client()
        driver.close()
        return yq, username, password, master_dict,types,reference_list
 
def getReports(num,chunk):
    print(f'Hello from {num}')
    print('Working on chunk: '+str(num))
    time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
    yr = time_stuff.iloc[0,0]
    qtr = time_stuff.iloc[0,1]
    yq=str(yr)+str(qtr)
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Vermont', usecols='A,B',dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    os.chdir('C:/Users/')
    chromeOptions = webdriver.ChromeOptions()
    prefs = {'download.default_directory':'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder\\Illinois\\',
         'plugins.always_open_pdf_externally':True,
         'download.prompt_for_download':False}
    chromeOptions.add_experimental_option('prefs',prefs)
    chromeOptions.add_argument('--disable-gpu')
    driver = webdriver.Chrome(chrome_options = chromeOptions, executable_path=r'C:\chromedriver.exe')
    os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder\\Vermont\\')
    driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
    user = driver.find_element_by_id('username')
    user.send_keys(username)
    pass_word = driver.find_element_by_id('password')
    pass_word.send_keys(password)
    login = driver.find_element_by_id('submit')
    in_flag = 0
    wait = WebDriverWait(driver,10)
    while in_flag == 0:
        login.click()
        try:
            canary = wait.until(EC.element_to_be_clickable((By.ID,'terms')))        
            in_flag = 1
        except TimeoutException as ex:
            driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
            user = driver.find_element_by_id('username')
            user.send_keys(username)
            pass_word = driver.find_element_by_id('password')
            pass_word.send_keys(password)
            login = driver.find_element_by_id('submit')
    
    
    accept = wait.until(EC.element_to_be_clickable((By.ID,'terms')))
    accept.click()
    invoices = driver.find_element_by_xpath('//a[text()="Invoices"]')
    invoices.click()
    type_dropdown = lambda: WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.ID,'docType')))
    type_select = lambda: Select(type_dropdown())
    types = type_select().options
    types = [item.text for item in types[1:]]
    types_2=[]
    for i,typ in enumerate(types):
        if len(typ.split(' '))==1:
            types_2.append(typ)
        elif len(typ.split('(')[-1].split(' '))==1:
            _ = typ.split(' ')[-1].replace('(','').replace(')','').replace(' ','_')
            types_2.append(_)
        elif len(typ.split(' '))==2:
            _ = typ.replace(' ','_')
            types_2.append(_)
        else:
            _='_'.join(typ.split('(')[1].split(' ')).replace(')','')
            types_2.append(_)
    reports_tab = wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@href="/RebateServicesPortal/reports/index"]')))       
    reports_tab.click()                
    report = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="reportList"]')))
    report_select = lambda: Select(report())
    #Now starting iterating through the chunk
    #for chunk in chunks:
    for label, program, ndc in chunk:
        success = 0
        if program == 'Medicare_Wrap':
            program = 'VPharm/SPAP (Medicare Wrap)'
        elif program == 'DME':
            program = 'State Only Diabetic (DME)'
        while success==0:
            try:
                report = driver.find_element_by_xpath('//select[@name="stateReportId"]')
                select_report = lambda: Select(report)
                if program == 'JCode':
                    select_report().select_by_index(2)
                    print('1')
                else:
                    select_report().select_by_index(1)
                    print('2')
                ndc_in = driver.find_element_by_xpath('//input[@name="ndc"]')
                print('a')
                ndc_in.send_keys(ndc)
                wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@name="docType"]')))
                docType = driver.find_element_by_xpath('//select[@name="docType"]')
                select_docType = Select(docType)
                select_docType.select_by_visible_text(program.replace('_',' '))
                wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@name="rpuStart"]')))
                rpu = driver.find_element_by_xpath('//input[@name="rpuStart"]')
                rpu.send_keys(yq)
                print('b')
                submit_button= driver.find_element_by_xpath('//input[@value="Submit"]')
                submit_button.click()
                accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Accept"]')))
                accept.click()
                wait.until(EC.staleness_of(accept))
                wait.until(EC.staleness_of(submit_button))
                soup = BeautifulSoup(driver.page_source,'html.parser')
                Reports = [x.text.strip() for x in soup.find_all('td')]
                report_type = types_2[types.index(program)]
                if program == 'JCode':
                    key = f'EXT JCODE CLD Report for {ndc} VT {yq} JCode'
                else:
                    key = f'EXT Claim Level Detail Report Report for {ndc} VT {yq} {report_type}'
                print('c')
                if any(map((lambda x: key in x),Reports)):
                    success=1
                    print('d')
                else:
                    driver.refresh()
                    print('e')
                    pass
            except TimeoutException as ex:
                time.sleep(10)
                driver.refresh()
                pass
            except NoSuchElementException as ex:
                time.sleep(10)
                driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
                user = driver.find_element_by_id('username')
                user.send_keys(username)
                pass_word = driver.find_element_by_id('password')
                pass_word.send_keys(password)
                login = driver.find_element_by_id('submit')
                login.click()
                reports_tab = wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@href="/RebateServicesPortal/reports/index"]')))       
                reports_tab.click()                
                print(ex)
                print('Trying again')
    driver.close()
    
									
    
def make_chunks(master_dict):
    #Break the information for each report down into 
    reports = []
    for key in master_dict.keys():
        for key2 in master_dict[key].keys():
            for value in master_dict[key][key2]:
                    report = (key,key2,value)
                    reports.append(report)
    import math
    n = math.ceil(len(reports)/3)
    chunks = [reports[x:x+n] for x in range(0,len(reports),n)]
    return chunks

##reports stuff is below this

  
    
def download_reports():
    os.chdir('O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder\\Vermont')
    time_stuff = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx', sheet_name = 'Year-Qtr',use_cols='A:B')
    yr = time_stuff.iloc[0,0]
    qtr = time_stuff.iloc[0,1]
    yq=str(yr)+str(qtr)
    chromeOptions = webdriver.ChromeOptions()
    prefs = {'download.default_directory':f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder\\Vermont\\',
             'plugins.always_open_pdf_externally':True}
    chromeOptions.add_experimental_option('prefs',prefs)
    driver =webdriver.Chrome(options = chromeOptions, executable_path=r'C:\chromedriver.exe')
    login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Vermont', usecols=[0,1],dtype='str')
    username = login_credentials.iloc[0,0]
    password = login_credentials.iloc[0,1]
    driver.get(r'https://www.vermontrsp.com/RebateServicesPortal/login/home?goto=http://www.vermontrsp.com/RebateServicesPortal/')
    user = driver.find_element_by_id('username')
    user.send_keys(username)
    pass_word = driver.find_element_by_id('password')
    pass_word.send_keys(password)
    login = driver.find_element_by_id('submit')
    login.click()
    
    wait = WebDriverWait(driver,10)
    accept = wait.until(EC.element_to_be_clickable((By.ID,'terms')))
    accept.click()
    reports_tab = driver.find_element_by_xpath('//a[text()="Reports"]')       
    reports_tab.click()                
    report = lambda: driver.find_element_by_xpath('//select[@id="reportList"]')
    report_select = lambda: Select(report())
    report_select().select_by_index(1)
    time.sleep(2)
    types = driver.find_element_by_xpath('//select[@id="docType"]')
    types_select = Select(types)
    programs = [x.text for x in types_select.options][1:]
    types_2=[]
    for i,typ in enumerate(programs):
        if len(typ.split(' '))==1:
            types_2.append(typ)
        elif len(typ.split('(')[-1].split(' '))==1:
            _ = typ.split(' ')[-1].replace('(','').replace(')','')
            types_2.append(_)
        elif len(typ.split(' '))==2:
            _ = typ.replace(' ','_')
            types_2.append(_)
        else:
            _='_'.join(typ.split('(')[1].split(' ')).replace(')','')
            types_2.append(_)
    values = [x.get_attribute('value') for x in driver.find_elements_by_xpath('//select[@id="docType"]/option')][1:]
    mapper = dict(zip(types_2,values))    
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
    links = [x.find_element_by_xpath('td//a[@href="#"]/i') for x in rows]
    master_df = pd.DataFrame()
    files = []
    for name, link in zip(names, links):
        success_flag = 0
        while success_flag==0:
            #get info for file name
            ndc = name.split(' ')[7]
            state = name.split(' ')[8]
            program = name.split(' ')[10]
            value = mapper[program]
            first_half = '_'.join(name.split(' ')[:5])
            second_half = '-'.join(name.split(' ')[-4:]).replace(program,mapper[program])
            download_name = '-'.join([first_half,second_half])+'.xls'
            if download_name in os.listdir():
                continue
            else:
                pass
            files.append(download_name)
            #download the file
            
            counter = 0
            try:
                link.click()
            except WebDriverException as ex:
                driver.refresh()
                time.sleep(10)
                continue
            while download_name not in os.listdir() and counter<10:
                time.sleep(1)
                counter+=1
            if counter >9:
                pass
            else:
                success_flag=1
    
        temp_df = pd.read_excel(download_name,skipfooter=3)
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
        try:
            program_code_name = mapper[splitter]
        except:
            program_code_name = splitter
        path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Vermont\\'+splitter+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
        file_name = 'VT_'+program_code_name+'_'+str(qtr)+'Q'+str(yr)+'.xlsx'
        if os.path.exists(path)==False:
            os.makedirs(path)
        else:
            pass
        os.chdir(path)
        frame.to_excel(file_name, engine='xlsxwriter', index=False)
    closers = lambda: driver.find_elements_by_xpath('//a[@title="Delete"]/i')
    for close in closers():
        closers()[0].click()
        alert = driver.switch_to.alert
        alert.accept()
    driver.stop_client()
    driver.close()
    for file in os.listdir():
        os.remove(file)



def multi_grabber(chunks):             
    processes = [mp.Process(target=getReports,args=(i,chunk)) for i,chunk in enumerate(chunks)]
    for p in processes:
        p.start()       
    for p in processes:
        p.join() 
        
        
def main():
   
    grabber = Vermont_Grabloid()
    grabber.delete_old_reports()
    yq, username, password, master_dict,types,reference_list = grabber.get_invoices()  
    chunks = make_chunks(master_dict)
    #getReports(1,chunks)
    multi_grabber(chunks)
    download_reports()

    
if __name__=='__main__':
    main()










