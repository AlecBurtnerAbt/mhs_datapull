# -*- coding: utf-8 -*-
"""
Created on Mon Nov 19 11:53:57 2018

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
import pandas as pd
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoAlertPresentException, InvalidElementStateException
from bs4 import BeautifulSoup
import gzip
import shutil
from grabloid import Grabloid
from tqdm import tqdm

class KansasGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script="Kansas")
        
    def pull(self):
        driver = self.driver
        qtr = self.qtr
        yr = self.yr
        username = self.credentials.iloc[0,0]
        password = self.credentials.iloc[0,1]
        driver.get('https://www.kmap-state-ks.us/provider/security/logon.asp')
        yq=str(yr)+str(qtr)
        wait = self.wait
        mapper= pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Delaware', usecols='A,B',dtype='str')
        mapper = dict(zip(mapper.iloc[:,0].str.upper(),mapper.iloc[:,1]))
        #Once page has loaded identify username, pw, and logon button. 
        #input required credentials and log in
        username_input = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@id="logonid"]')))
        username_input.send_keys(username)
        password_input = driver.find_element_by_xpath('//input[@id="logonpswd"]')
        password_input.send_keys(password)
        logon_button = driver.find_element_by_xpath('//input[@name="submit2"]')
        logon_button.click()
        
        #page comes up with disclaimer notice, "Next" button must be clicked
        next_button = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@name="nextButton"]')))
        next_button.click()
        
        #presented with a page of files whose names are download links
        #so, I'll scrape the names that have the target year/qtr in them
        #and use the file name to meet the naming convention
        
        file_names = driver.find_elements_by_xpath('//tbody[//font[text()="Filename"]]//a')
        file_links = [x for x in file_names if yq in x.text]
        file_names = [x.text for x in file_names if yq in x.text]
        state = 'KS'
        state_full = 'Kansas'
        #KS is a by NDC state, so initialize an empty dict
        cld_to_get = []
        
        
        for link, name in tqdm(zip(file_links, file_names)):
            success = 0
            #pick out elements to build correct file name
            labeler = name.split('_')[2]
            website_program = name.split('_')[3].split('.')[0]
            lilly_program = mapper[website_program]
            ext = name[-4:]
            while success == 0:
                link.click()
                print(f'Link clicked for {name}')
                counter = 0
                while name not in os.listdir() and counter<10:
                    print(f'Waiting for {name}')
                    counter +=1
                    time.sleep(1*counter)
                if name not in os.listdir():
                    continue
                else:
                    success=1
            read_success = 0
            counter = 0
            if ext == '.txt':
                while read_success == 0:
                    try:
                        with open(name) as F:
                            lines = F.readlines()
                            ndcs = [x[6:17] for x in lines]
                            ndcs = list(set(ndcs))
                            for ndc in ndcs:
                                cld_to_get.append((labeler,website_program,ndc))
                            read_success=1
                    except:
                        print(f'Read failed for {os.getcwd()} {name}, retrying')
                        time.sleep(2)
                        counter +=1
                        if counter >9:
                            break
                        
            else:
                pass
            #Build the correct name
            file_name = f'{state}_{lilly_program}_{qtr}Q{yr}_{labeler}{ext}'
            invoice_directory = f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\{state_full}\\{lilly_program}\\{yr}\\Q{qtr}\\'
            if os.path.exists(invoice_directory) == False:
                os.makedirs(invoice_directory)
            else:
                pass
            shutil.move(name,invoice_directory+file_name)            
        
        #Moving onto CLD
        cld_link = driver.find_element_by_xpath('//a[text()="Claim Level Detail"]')
        cld_link.click()  
        
        #Accept the dislcaimer
        accept_button = driver.find_element_by_xpath('//input[@value="Accept"]')          
        accept_button.click()            
        
        for item in cld_to_get:
            labeler = item[0]
            program = item[1]
            ndc = item[2]
            time_box = driver.find_element_by_xpath('//input[@id="input_Quarter"]')
            time_box.send_keys(yq)
            ndc_box = driver.find_element_by_xpath('//input[@id="input_NDC"]')
            ndc_box.send_keys(ndc)
            program_select = driver.find_element_by_xpath('//select[@id="input_InvoiceType"]')
            program_select = Select(program_select)            
            program_select.select_by_value(program)
            search_button = driver.find_element_by_xpath('//button[@id="submit_button"]')
            search_button.click()



a = KansasGrabloid()            
yr = 2018            
qtr=3            
driver = a.driver            
username = a.credentials.iloc[0,0]
password = a.credentials.iloc[0,1]            
wait = a.wait      
            
            
            
            
            
            
            
            
            
            
            
            