# -*- coding: utf-8 -*-
"""
Created on Wed Aug 29 16:06:43 2018

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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoAlertPresentException
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
from grabloid import Grabloid

class NorthDakotaGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='North Dakota')

    def pull(self):
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait    
        login_credentials = self.credentials
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='North Dakota', usecols='D,E',dtype='str')
        mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        #Login with provided credentials
        driver.get('https://drugrebate.nd.gov/RebateWeb/login.do')
        wait2 = WebDriverWait(driver,3)
        user =  wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@id="username"]')))  
        user.send_keys(username)
        pw = driver.find_element_by_xpath('//input[@id="password"]')
        pw.send_keys(password)
        login = driver.find_element_by_xpath('//input[@alt="Login"]')
        login.click()
        
        #Now navigate to the invoices
        
        invoices = wait.until(EC.element_to_be_clickable((By.XPATH,'//img[@alt="Search Invoice"]')))
        invoices.click()
        
        
        #Have to enter all required fields
        quarter = wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="startQq"]')))
        quarter.send_keys(str(qtr))
        
        year = driver.find_element_by_xpath('//input[@id="qtrYear"]')
        year.send_keys(str(yr))
        
        programs = lambda: driver.find_element_by_xpath('//select[@id="rebateProgram"]')
        program_select = lambda: Select(programs())
        progs = [x.text for x in program_select().options if len(x.text)>1]
        
        payer = driver.find_element_by_xpath('//select[@name="payer"]')
        payer.send_keys('N')
        
        status = driver.find_element_by_xpath('//select[@name="downloadStatus"]')
        status.send_keys('B')
        
        form = lambda: driver.find_element_by_xpath('//select[@id="fileFormat"]')
        form_select = lambda: Select(form())
        form_types = [x.text for x in form_select().options if len(x.text)>1]
        form_types = form_types[2:4]
        
        
        labels = lambda: driver.find_element_by_xpath('//select[@id="labelerCode"]')
        label_select = lambda: Select(labels())
        labelers = [x.text for x in label_select().options if len(x.text)>1]
        
        
        submit_button = lambda: driver.find_element_by_xpath('//img[@alt="Submit form"]')
        
        #create empty list far email notification
        invoices_obtained=[]
        #Now that all Selects have been assigned to functions begin looping
        
        for F in form_types:
            form_select().select_by_visible_text(F)
            submit_button().click()
            check_all = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="all"]')))
            check_all.click()
            download_link = driver.find_element_by_xpath('//img[@name="download"]')
            download_link.click()
            try:
                alert = driver.switch_to.alert
                alert.accept()
                time.sleep(3)
            except NoAlertPresentException as ex:
                pass
            while any(map((lambda x: 'zip' in x),os.listdir()))==False:
                time.sleep(1)
            file = os.listdir()[0]
            with zipfile.ZipFile(file) as ax:
                ax.extractall()
            os.remove(file)
            soup = BeautifulSoup(driver.page_source,'html.parser')
            table = soup.find('table',attrs={'id':'inv'})
            rows = table.find_all('tr')[1:]
            raw = []
            [raw.append(x.text.strip().replace('\n','_')) for x in rows]
            P = [x.split('_')[3] for x in raw]
            L = [x.split('_')[4] for x in raw]
            file_names = ['ND_{}_{}Q{}_{}'.format(mapper[x],qtr,yr,y) for x,y in zip(P,L)] 
            #using the html tags I'll make a dictionary to help name files
            keys = [x.split('_')[0] for x in raw]
            mapp = dict(zip(keys,file_names))
            for file in os.listdir():
                file_name = mapp[file.split('.')[0]]
                if F =='Print Image':
                    extension = '.pdf'
                else:
                    extension = '.txt'
                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\North Dakota\\'+ mapp[file.split('.')[0]].split('_')[1]+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                if os.path.exists(path)==False:
                    os.makedirs(path)
                else:
                    pass
                invoices_obtained.append(file_name)
                shutil.move(file,path+file_name+extension)
            driver.back()
        
        driver.close()
        
        return invoices_obtained
    
def main():
    grabber = NorthDakotaGrabloid()
    invoices = grabber.pull()
    grabber.send_message(invoices)

if __name__=='__main__':
    main()    
    
























