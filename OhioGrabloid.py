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
import xlsxwriter as xl
from grabloid import Grabloid, push_note


class OhioGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Ohio')



    def pull():
        driver = self.driver
        yr = self.yr
        qtr = self.qtr
        login_credentials = self.credentials
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        yq=str(yr)+str(qtr)
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Ohio', usecols='D,E',dtype='str')
        mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
        #Login with provided credentials
        driver.get('https://rsp.ohgov.changehealthcare.com/RebateServicesPortal/login/home?goto=http://rsp.ohgov.changehealthcare.com/RebateServicesPortal/')   
        
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
        invoices = driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div/ul/li[2]/a')
        invoices.click()
        code_dropdown = lambda: driver.find_element_by_id('labeler')
        code_select = lambda: Select(code_dropdown())
        codes = [x.text for x in code_select().options][1:]
        type_dropdown = lambda: WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.ID,'docType')))
        type_select = lambda: Select(type_dropdown())
        master_dict = dict.fromkeys(codes)
        time_stamp = lambda: driver.find_element_by_xpath('//input[@name="period"]')
        types = [x.text for x in type_select().options][1:]
        invoices_obtained = []
        
        for code in codes:
            code_select().select_by_visible_text(code)
            print('Selecting '+str(code))
            report_dict = {}
            for report in types:
                wait.until(EC.presence_of_element_located((By.XPATH,'//select[@id="docType"]')))
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
                    file = 'OH-'+code+'-'+yq+'-'+report+file_type
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
                        with open(file) as ax:
                            lines = ax.readlines()
                            ndcs = list(set([x[6:17] for x in lines]))
                            report_dict.update({report:ndcs})
                    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Ohio\\'+report+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                    if os.path.exists(path)==False:
                        os.makedirs(path)
                    else:
                        pass
                    file_name = 'OH_{}_{}Q{}_{}{}'.format(report,qtr,yr,code,file_type)
                    invoices_obtained.append(file_name)
                    shutil.move(file,path+file_name)
            master_dict.update({code:report_dict})
        
             #############################################CLD below
        #This only requests the generation of the CLD reports, a new function has to 
        #be created to go back later and download those reports
        reports_tab = driver.find_element_by_xpath('//a[text()="Reports"]')       
        reports_tab.click()                
        report = lambda: driver.find_element_by_xpath('//select[@id="reportList"]')
        report_select = lambda: Select(report())
        not_ready = []
        
        for labeler in list(master_dict.keys()):
            master_df = pd.DataFrame()
            for program in list(master_dict[labeler].keys()):
        
                for ndc in master_dict[labeler][program]:
                    report_select().select_by_index(1)
                    time.sleep(.5)
                    st_prog = program.replace('_',' ')
                    try:
                        lly_prog = mapper[st_prog]
                    except KeyError as err:
                        lly_prog = st_prog
                    reports = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="docType"]')))
                    reports_select = lambda: Select(reports())
                    reports_select().select_by_visible_text(st_prog)
                    if len(ndc)<2:
                        continue
                    else:
                        pass
                    ndc_in = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@name="ndc"]')))
                    ndc_in.send_keys(ndc)
                    time_stamp = driver.find_element_by_xpath('//input[@name="rpuStart"]')
                    time_stamp.send_keys(yq)
                    time_stamp2 = driver.find_element_by_xpath('//input[@name="rpuEnd"]')
                    time_stamp2.send_keys(yq)
                    submit_button= driver.find_element_by_xpath('//input[@value="Submit"]')
                    submit_button.click()
                    accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Accept"]')))
                    accept.click()
                    wait.until(EC.staleness_of(accept))
                    time.sleep(1)



    def download_reports(self):
        driver = self.driver
        wait = self.wait
        qtr = self.qtr
        yr = self.yr
        login_credentials = self.credentials
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        yq=str(yr)+str(qtr)
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Ohio', usecols='D,E',dtype='str')
        mapper = dict(zip(mapper['State Program'],mapper['Flex Program']))
        driver.get('https://rsp.ohgov.changehealthcare.com/RebateServicesPortal/login/home?goto=http://rsp.ohgov.changehealthcare.com/RebateServicesPortal/')   

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
        
        #At the landing page after login and accepting terms of use
        #need to navigate to reports page 
        
        reports_link = driver.find_element_by_xpath('//a[contains(text(),"Reports")]')
        reports_link.click()        
                
def main():
    grabber = OhioGrabloid()
    
    
if __name__=='__main__':
    main()



yr, qtr, login_credentials = grabber.yr, grabber.qtr, grabber.credentials
driver = grabber.driver

