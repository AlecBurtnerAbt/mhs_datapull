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
        files = lambda: driver.find_elements_by_xpath(f'//td[contains(text(),{yq})]')
        for j,file in enumerate(files()):
            file_to_get = files()[j]
            name = file_to_get.text
            files()[j].click()
            time.sleep(10)
            wait.until(EC.staleness_of((file)))
            labeler_code = name[1:6]
            two_letter_code = name[-6:-4]
            program = two_letter_mapper[two_letter_code]
            download_button_values = [ x.get_attribute('value') for x in driver.find_elements_by_xpath('//input[contains(@value,"Download")]')]
            # have the buttons to click to download files
            for value in download_button_values:    
                button_to_click = driver.find_element_by_xpath(f'//input[@value="{value}"]')
                wait.until(EC.element_to_be_clickable((By.XPATH,f'//input[@value="{value}"]')))
                button_to_click.send_keys(Keys.RETURN)
                if 'PDF' not in button_to_click.get_attribute('value'):
                    while len(os.listdir()) <1:
                        print('Waiting on file to download...')
                        time.sleep(1)
                    while any(map((lambda x: '.crd' in x),os.listdir())) or any(map((lambda x: '.tmp' in x),os.listdir())):
                        print('Waiting for download to finish')
                        time.sleep(1)
                    ext = '.txt'
                else:
                    #time.sleep(8)
                    while len(driver.window_handles) <2:
                        print('waiting for other window to open')
                        time.sleep(2)
                    driver.switch_to.window(driver.window_handles[-1])
                    frame = wait.until(EC.presence_of_element_located((By.NAME,'frmViewer')))
                    driver.switch_to.frame(frame)
                    open_button = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="open-button"]')))
                    ActionChains(driver).move_to_element(open_button).click().perform()
                    while len(os.listdir()) <1:
                        print('Waiting on file to download...')
                        time.sleep(1)
                    while any(map((lambda x: '.crd' in x),os.listdir())) or any(map((lambda x: '.tmp' in x),os.listdir())):
                        print('Waiting for download to finish')
                        time.sleep(1)
                    ext = '.pdf'
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                download_file = os.listdir()[0]
                file_name = f'WI_{program}_{qtr}Q{yr}_{labeler_code}_{ext}'
                path = f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\Wisconsin\\{program}\\{yr}\\Q{qtr}\\'
                if os.path.exists(path) == False:
                    os.makedirs(path)
                shutil.move(download_file,os.path.join(path,file_name))
                print(f'Done with {file_name}')
        #Now we move on to grab the CLD 
        
        manufacturer_link = driver.find_element_by_xpath('//a[text()="Manufacturer"]')
        manufacturer_link.click()                

        # Now click on CLD requests
        #navigate to cld section        
        cld_link = wait.until(EC.element_to_be_clickable((By.XPATH,'//a[@title="Claim Level Detail (CLD) Requests"]')))
        cld_link.click()
    
        #select the option to download CLD for an entire invoice
        cld_for_whole_invoice_link = driver.find_element_by_xpath('//a[@title="Request CLD for an Entire Invoice"]')
        cld_for_whole_invoice_link.click()

        #There are drop downs for labeler code and invoice type,
        # and an input box for the year quart in Q/YYYY format
        # So we'll get the values in the drops downs first, then
        # loop through the options requesting CLD data
        
        labeler_code_drop_down = lambda: driver.find_element_by_xpath('//select[contains(@id,"LabelerCode")]')
        labeler_code_select = lambda: Select(labeler_code_drop_down())
        labeler_codes = [option.text for option in labeler_code_select().options[1:]] #first option is blank space, cut it out
        
        invoice_type_drop_down = lambda: driver.find_element_by_xpath('//select[contains(@name,"InvoiceType")]')
        invoice_type_select = lambda: Select(invoice_type_drop_down())        
        invoices = [option.text for option in invoice_type_select().options[1:]]#again the first option is blank space, cut it out
        
        date_input = lambda: driver.find_element_by_xpath('//input[contains(@name,"InvoicePeriod")]')
        #going to keep it to a single FOR loop by using product of 
        # all invoices and labeler codes
        # provides generator that returns tuples of (labeler, invoice)
        all_options = product(labeler_codes, invoices)
        
        for labeler, invoice in all_options:
            try:
                print(f'Finding CLD for {labeler.strip()}:{invoice.strip()}')
                labeler_code_select().select_by_visible_text(labeler)
                date_input().send_keys(f'{qtr}/{yr}')        
                invoice_type_select().select_by_visible_text(invoice)            
                submit_button = driver.find_element_by_xpath('//input[@value="Submit"]')
                submit_button.send_keys(Keys.RETURN)
                wait.until(EC.staleness_of(submit_button))   
                #after submit button is clicked it takes you to another page, so
                #navigate back to the CLD page
                cld_home_link = driver.find_element_by_xpath('//a[text()="Claim Level Detail Home"]')
                cld_home_link.click()
                cld_for_whole_invoice_link = driver.find_element_by_xpath('//a[@title="Request CLD for an Entire Invoice"]')
                cld_for_whole_invoice_link.click()
            except TimeoutException as ex:
                print(ex)
                print('Timed out')
            try:
                no_data_present_message = driver.find_element_by_xpath('//a[contains(text(),"We did not find any claims")]')
                print(f'No data present for {labeler.strip()}:{invoice.strip()}')
            except:
                print('No errors')
            finally:
                print(f'Done with {labeler.strip()}:{invoice.strip()}')
                
     print('Requesting CLD complete, closing browser')           
    '''
    It takes 24 hours to generate reports, run
    WisconsinGrabloid2 the next day
    '''




@push_note(__file__)
def main():

    grabber = WisconsinGrabloid()
    invoices = grabber.pull()
    grabber.send_message(invoices)
    
if __name__=='__main__':
    main()



