# -*- coding: utf-8 -*-
"""
Created on Tue Jun  4 10:05:36 2019

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

class RhodeIslandGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Rhode Island')
        script = self.script
        self.mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Rhode Island', usecols='D,E',dtype='str')
        
    def pull(self):
        driver = self.driver
        yr = str(self.yr)
        yr2 = str(yr)[-2:]
        qtr = str(self.qtr)
        program_mapper = self.mapper
        program_mapper = dict(zip(program_mapper['Portal ID'],program_mapper['Flex ID']))
        login_credentials = self.credentials
        wait = self.wait
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        driver.get('https://ridhssecureftp.com/EFTClient/Account/Login.htm')
        #Now that we're at the page login
        user_name = wait.until(EC.element_to_be_clickable((By.ID,'username')))
        user_name.send_keys(username)
        pass_word = driver.find_element_by_id('password')
        pass_word.send_keys(password)
        submit_button = driver.find_element_by_id('loginSubmit')
        submit_button.click()
        
        #Okay, logged in.   
        # All the 00002 files are on the page after log in, the other labeler codes are in folders
        # so we'll get the current quarter 00002 files and then move onto the other labeler codes
        
        files = [x for x in driver.find_elements_by_class_name('ng-binding') if x.text[:3] == f'{yr2}{qtr}']
        labeler = '00002'
        for file in files:
            file.click()
            while file.text not in os.listdir():
                time.sleep(1)
                print(f'Waiting for {file.text} to download')
            name = file.text
            flex_code = program_mapper[name[3]]
            extension = name[-3:]
            file_name = f'RI_{flex_code}_{qtr}Q{yr}_{labeler}.{extension}'
            path = f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Rhode Island\\{flex_code}\\{yr}\\Q{qtr}\\'
            if os.path.exists(path)==False:
                os.makedirs(path)
            shutil.move(name,os.path.join(path,file_name))
        
        #now that we have files for 00002 we'll move to the other labelers
        labeler_codes = ['66733','00777']        
        for labeler in labeler_codes:
            print('Finding folder')
            folder = wait.until(EC.presence_of_element_located((By.XPATH,f'//*[@id="drop-zone-container"]//a[@class="ng-binding" and contains(text(),"DR{labeler}")]')))
            print('Clicking folder')
            folder.click()
            time.sleep(10)
            files = [x for x in driver.find_elements_by_class_name('ng-binding') if x.text[:3] == f'{yr2}{qtr}']
            print(f'Getting {len(files)} files')
            for file in files:
                file.click()
                while file.text not in os.listdir():
                    time.sleep(1)
                    print(f'Waiting for {file.text} to download')
                name = file.text
                flex_code = program_mapper[name[3]]
                extension = name[-3:]
                file_name = f'RI_{flex_code}_{qtr}Q{yr}_{labeler}.{extension}'
                path = f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Rhode Island\\{flex_code}\\{yr}\\Q{qtr}\\'
                if os.path.exists(path)==False:
                    os.makedirs(path)
                shutil.move(name,os.path.join(path,file_name))
            print('Going to parent folder')
            driver.back()
            
        driver.close()
            
if __name__ == '__main__':
    grabber = RhodeIslandGrabloid()
    grabber.pull()            
            
            
            
            

grabber = RhodeIslandGrabloid()
yr = grabber.yr
qtr = grabber.qtr
login_credentials = grabber.credentials
driver = grabber.driver
program_mapper = grabber.mapper
wait = grabber.wait
//*[@id=]/div/div/list-view/div/ul/li[3]/div[1]/span/div[2]/a
<span ng-show="row.branch.data.isHeader" ></span>