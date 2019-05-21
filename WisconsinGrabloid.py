# -*- coding: utf-8 -*-
"""
Created on Tue May 21 15:10:02 2019

@author: AUTOBOTS
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
import pickle
from itertools import product
import logging

class WisconsinGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Wisconsin')

    
    def pull(self, efficient=True):
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        login_credentials = self.credentials
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        yq = str(yr)+str(qtr)
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Wisconsin', usecols='G,E',dtype='str')
        mapper['two_letter_code'] = mapper['WI Portal Program Name'].str[:2]
        two_letter_mapper = dict(zip(mapper[list(mapper.columns)[2]],mapper[list(mapper.columns)[0]]))
        #Login with provided credentials
        driver.get('https://www.forwardhealth.wi.gov/WIPortal/Default.aspx ')   
        drug_rebate_button = driver.find_element_by_xpath('//a[contains(@href,"DrugRebateLogin")]')
        drug_rebate_button.click()
        username_input = wait.until(EC.presence_of_element_located((By.XPATH,'//input[contains(@name,"userName")]')))
        username_input.send_keys(username)
        password_input = driver.find_element_by_xpath('//input[contains(@name,"password")]')
        password_input.send_keys(password)
        login_button = driver.find_element_by_xpath('//a[contains(@id,"LoginButton2")]')
        login_button.click()
        #logged into the system, now have to get invoices
        invoice_link = wait.until(EC.presence_of_element_located((By.XPATH,'//a[contains(text(),"Download Invoices")]')))
        invoice_link.click()
        
        #takes you to a page where the files are
        table = wait.until(EC.presence_of_element_located((By.XPATH,'//table[contains(@id,"Datalist")]')))
        files = table.find_elements_by_xpath(f'//td[contains(text(),{yq})]')
        for file in files:
            name = file.text
            labeler_code = name[1:6]
            two_letter_code = name[-6:-4]
            program = two_letter_mapper[two_letter_code]
            file_name = f'WI_{program}_{qtr}Q{yr}_'
            download_buttons = driver.find_elements_by_xpath('//input[contains(@value,"Download")]')
            # have the buttons to click to download files
            for button in download_buttons():            
                button.click()
                while len(os.listdir()) <1:
                    print('Waiting on file to download...')
                    time.sleep(1)
                while '.txt' in os.listdir()
            
@push_note(__file__)
def main():

    grabber = WisconsinGrabloid()
    invoices = grabber.pull()
    grabber.send_message(invoices)
    
if __name__=='__main__':
    main()



