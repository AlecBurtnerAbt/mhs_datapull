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
from grabloid import Grabloid

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
        for label in options:
            labeler_select().select_by_visible_text(label)
            for report in list(mapper.keys()):
                submit_button = driver.find_element_by_xpath('//input[@value="Submit"]')
                types_select().select_by_visible_text(report)
                time_stamp().send_keys(yq)
                submit_button.click()
                report_value = driver.find_element_by_xpath('//select[@id="docType"]/option[text()="{}"]'.format(report)).get_attribute('value')
                while wait.until(EC.staleness_of(submit_button))==False:
                    time.sleep(.2)
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
                    while file_name in os.listdir()==False:
                        time.sleep(1)
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
                            
        
        driver.close()
        return invoices_obtained
def main():
    grabber = IllinoisGrabloid()
    invoices = grabber.pull()
    grabber.send_message(invoices)
    
if __name__=='__main__':
    main()














