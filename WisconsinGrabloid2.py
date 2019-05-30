# -*- coding: utf-8 -*-
"""
Created on Mon Apr 29 10:01:05 2019

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
from itertools import product
from tqdm import tqdm
import pickle


class WisconsinGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Wisconsin')

    
    def pull(self, efficient=True):
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        login_credentials = self.credentials
        #make sure the directory is the downloads folder!
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Wisconsin', usecols='E,G',dtype='str')
        mapper = dict(zip(mapper['WI Portal Program Name'],mapper['MRB Contract ID']))
        #Login with provided credentials
        driver.get('https://www.forwardhealth.wi.gov/WIPortal/Subsystem/DrugRebate/DrugRebateLogin.aspx') 
        user_name = driver.find_element_by_xpath('//input[contains(@id,"userName")]')
        user_name.send_keys(username)
        pass_word = driver.find_element_by_xpath('//input[contains(@id,"password")]')
        pass_word.send_keys(password)
        login_button = driver.find_element_by_xpath('//a[contains(@id,"LoginButton2")]')
        login_button.click()
        #okay, logged in , click on the invoices link
        '''
        There were no invoices available on the website as of 4/25/2019, dealying development of 
        invoice section until invoices are present to determine structure
        invoice_link = driver.find_element_by_xpath('//a[@title="Download Invoices"]')
        invoice_link.click()
        '''
        #navigate to cld section        
        cld_link = wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@title="Claim Level Detail (CLD) Requests"]')))
        cld_link.click()
        
        #Now we have to click on the "Download CLD Requests" link
        download_cld_requests_link = driver.find_element_by_xpath('//a[text()="Download CLD Requests"]')
        download_cld_requests_link.click()
        
        #Okay, the invoice identifiers are only two letter codes in the middle of the file name, so 
        # what we have to do is change up our mapper dictionary a little bit
        
        mapper2 = {k[:2]:v for k,v in mapper.items()}
        mapper3 = {k:v for k,v in zip(mapper2.keys(),mapper.keys())}#in case mapper fails use to pull back full name of program
        #Now we can identify the invoices available, I'll use 
        # pandas to get the link text and then iterate through that
        wait.until(EC.staleness_of(download_cld_requests_link))
        wait.until(EC.presence_of_element_located((By.XPATH,'//table[contains(@id,"NavItemLoadedControl")]')))
        raw_table = driver.find_element_by_xpath('//table[contains(@id,"NavItemLoadedControl")]')
        link_table = pd.read_html(raw_table.get_attribute('innerHTML'))[-1]
        link_table = link_table.dropna(axis=1).rename(columns=link_table.iloc[0]).drop(0)
        
        #begin looping through files
        for file in  tqdm(link_table['File Name']):
            success_flag = 0
            while success_flag == 0:
                try:
                    print(f'\nWorking on {file}')
                    download_link = driver.find_element_by_xpath(f'//td[text()="{file}"]')
                    download_link.click()
                    while len(os.listdir()) == 0:
                        print('waiting for download...')
                        time.sleep(2)
                    while any('.crd' in file for file in os.listdir()) or any('.tmp' in file for file in os.listdir()):
                        print('waiting for full download...')
                        time.sleep(2)
                    print(f'Successfully downloaded {file}')
                    download = os.listdir()[0]
                    website_program_identifier = file.split('.')[1].split('_')[2]
                    print(f'Web id is {website_program_identifier}')
                    try:
                        lilly_flex_code = mapper2[website_program_identifier]
                    except KeyError:
                        lilly_flex_code = mapper3[website_program_identifier]
                    labeler_code = download.split('.')[1].split('_')[0]
                    print(f'Lilly code is {lilly_flex_code}')
                    new_file_name = f'WI_{lilly_flex_code}_{qtr}Q{yr}_{labeler_code}_.xlsx'
                    print(f'Reading in data for {file}')
                    file_data = pd.read_csv(download)     
                    os.remove(download)
                    path = f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Wisconsin\\{lilly_flex_code}\\{yr}\\Q{qtr}'
                    if os.path.exists(path)==False:
                        os.makedirs(path)
                    print(f'Writing data to file...')
                    file_data.to_excel(os.path.join(path,new_file_name),index=False)
                    print('Success!\n')
                    success_flag = 1
                except:
                    print(f'Some error, retrying')
                    time.sleep(5)
        driver.close()
            
            
            
def main():
    grabber = WisconsinGrabloid()
    grabber.pull()
    
if __name__ == "__main__":
    main()
'''
driver = grabber.driver
qtr = grabber.qtr
yr = grabber.yr
login_credentials = grabber.credentials
wait = grabber.wait
'''