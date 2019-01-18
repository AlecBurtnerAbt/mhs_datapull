# -*- coding: utf-8 -*-
"""
Created on Wed Nov  7 09:53:12 2018

@author: C252059
"""

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import time
import os
from win32com.client import Dispatch
import pandas as pd
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import gzip
import shutil
import pandas as pd
import pprint
from pushover_complete import PushoverAPI
import sys
import traceback
import logging

class Grabloid():
    
    def __init__(self,script,use_chrome=True):
        os.chdir('O:\\')
        self.script = script
        self.temp_folder_path = f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder\\{self.script}\\'
        driver_path = "O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Automation Scripts Parameters\\"
        chrome_options = Options()
        prefs = {'download.default_directory':self.temp_folder_path,
         'plugins.always_open_pdf_externally':True,
         'download.prompt_for_download':False}
        chrome_options.add_experimental_option("prefs",prefs)
        if use_chrome == True:
            self.driver = webdriver.Chrome(executable_path=driver_path+"chromedriver.exe", options=chrome_options)
        else: 
            pass
        self.yr = pd.read_excel(driver_path+"automation_parameters.xlsx",sheet_name="Year-Qtr",usecols="A").iloc[0,0]
        self.qtr = pd.read_excel(driver_path+"automation_parameters.xlsx",sheet_name="Year-Qtr",usecols="B").iloc[0,0]
        self.credentials = pd.read_excel(driver_path+"automation_parameters.xlsx", sheet_name= f"{self.script}" ,usecols="A:B")
        self.to_address = pd.read_excel(driver_path+"automation_parameters.xlsx", sheet_name= "Email Address" ,usecols="A").iloc[0,0]
        self.wait = WebDriverWait(self.driver,10)
        if os.path.exists(self.temp_folder_path)==False:
                os.mkdir(self.temp_folder_path)
        else:
            pass
        files = os.listdir(self.temp_folder_path)
        for file in files:
            os.remove(self.temp_folder_path+file)
        os.chdir(self.temp_folder_path)
        log_location = "O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Automation Log\\"
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        fh = logging.FileHandler(os.path.join(log_location,'AutoBots.log'))
        fh.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        logger.addHandler(fh)
        logging.warning(f"{os.getlogin()} has run script for {self.script}")

        
    def send_message(self,invoices):
        subject = f'{self.script} Invoices'
        body = 'The following invoices have been downloaded:\n'+'\n'.join(invoices)
        base = 0x0
        obj = Dispatch('Outlook.Application')
        newMail = obj.CreateItem(base)
        newMail.Subject = subject
        newMail.Body = body+'\n'+str(pprint.pformat(body2))
        newMail.To = recipient
        newMail.display()
        newMail.Send()
    
    def cleanup(self):
        os.chdir('O:\\')
        clean = 0
        counter = 0
        while clean ==0 and counter <10:
            try:
                os.removedirs(self.temp_folder_path)
            except:
                counter +=1
                time.sleep(counter*1)
                continue
        
        
        
        

def push_note(script):
    p = PushoverAPI('a4u1afrfsocorp6r1cdes1ydn5g2m6')
    def inner_push(func):
        
        '''
        separate function from Grabloid, used to notify Alec Burtner-Abt (primary dev)
        of script failure or success while running bots for CMA team.  Leverages Pushover App
        '''
        
        
        def func_wrapper(*args,**kwargs):
            try:
                func(*args, **kwargs)
                p.send_message('ukdn5gtjkaejnd6qmwy42ej2yofmsz', f'{script} has successfully run.')
            except:
                
                tb = sys.exc_info()[2]
                tb = traceback.print_tb(tb)
                err = sys.exc_info()[0]
                print(err)
                print (tb)
                p.send_message('ukdn5gtjkaejnd6qmwy42ej2yofmsz', f'{script} did not terminate properly.\n {traceback.print_tb(sys.exc_info()[2])}')
        return func_wrapper
    return inner_push


if __name__=='__main__':
    print('Ok')
        
        

