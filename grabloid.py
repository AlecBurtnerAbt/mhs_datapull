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



class Grabloid():
    
    def __init__(self,script):
        os.chdir('O:\\')
        self.script = script
        self.land_folder = f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Landing_Folder\\{self.script}\\'
        driver_path = "O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Automation Scripts Parameters\\"
        chrome_options = Options()
        prefs = {'download.default_directory':self.land_folder,
         'plugins.always_open_pdf_externally':True,
         'download.prompt_for_download':False}
        chrome_options.add_experimental_option("prefs",prefs)
        self.driver = webdriver.Chrome(executable_path=driver_path+"chromedriver.exe", options=chrome_options)
        self.yr = pd.read_excel(driver_path+"automation_parameters.xlsx",sheet_name="Year-Qtr",usecols="A").iloc[0,0]
        self.qtr = pd.read_excel(driver_path+"automation_parameters.xlsx",sheet_name="Year-Qtr",usecols="B").iloc[0,0]
        self.credentials = pd.read_excel(driver_path+"automation_parameters.xlsx", sheet_name= f"{self.script}" ,usecols="A:B")
        self.to_address = pd.read_excel(driver_path+"automation_parameters.xlsx", sheet_name= "Email Address" ,usecols="A").iloc[0,0]
        self.wait = WebDriverWait(self.driver,10)
        if os.path.exists(self.land_folder)==False:
                os.mkdir(self.land_folder)
        else:
            pass
        files = os.listdir(self.land_folder)
        for file in files:
            os.remove(self.land_folder+file)
        os.chdir(self.land_folder)
    
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
        

if __name__=='__main__':
    print('Ok')
        
        

