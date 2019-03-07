# -*- coding: utf-8 -*-
"""
Created on Mon Jul 30 08:45:46 2018

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
from grabloid import Grabloid

class PrimsGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Prims')




    def pull(self):
        yr = self.yr
        qtr = self.qtr
        login_credentials = self.credentials
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        driver = self.driver
        driver.implicitly_wait(30)
        wait = WebDriverWait(driver,15)
    
        driver.get('https://primsconnect.dxc.com')
        i_accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_RadCheckBoxAccept"]/span[1]')))
        i_accept.click()
        user_name = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtUserName"]')
        user_name.send_keys(username)
        pass_word = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtPassword"]')
        pass_word.send_keys(password)
        login = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_btnLogin_input"]')
        login.click()    
        yq = str(yr)+str(qtr)
        yq2 = 'Q{}-{}'.format(qtr,yr)
        yq3 = '{}-Q{}'.format(yr,qtr)
        #Now inside the webpage, begin selection process
        submit_request = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Submit Request"]')))
        submit_request.click()      
        
        #Now in the request page, navigate to invoice tab
        
        invoice_request_page = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//span[text()="Electronic Invoice (TXT)"]')))
        soup = BeautifulSoup(driver.page_source,'html.parser')
        lists = soup.find_all('ul',attrs={'class':'rcbList'})    
        states = [x.text for x in lists[0]]    
        pdf_request_page = lambda: driver.find_element_by_xpath('//span[text()="Paper Invoice (PDF)"]')
        pages = [invoice_request_page,pdf_request_page]
        retrieved = []
        for page in pages:
            try:
                reset_button = driver.find_element_by_xpath('//input[@value="Reset"]')
                page().click()
                wait.until(EC.staleness_of(reset_button))
                
                for state in states:
                    #first make sure the right state is selected
                    state_input = driver.find_element_by_xpath('//input[contains(@name,"StateDropDown")]')
                    if state_input.get_attribute('value')==state:
                        pass
                    else:
                        state_to_select = driver.find_element_by_xpath('//div[contains(@id,"StateDropDown")]//li[text()="'+state+'"]')
                        ActionChains(driver).move_to_element(state_input).click().pause(1).click(state_to_select).pause(1).perform()
                        wait.until(EC.staleness_of(state_input))
                    #now get all the programs for the drop downs
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    lists = soup.find_all('ul',attrs={'class':'rcbList'})
                    programs = [x.text for x in lists[1]]
                    
                    for program in programs:
                        #Make sure the right program is slected, if not select it
                        program_drop_down = driver.find_element_by_xpath('//input[contains(@id,"ProgramDropDown_Input")]')
                        if program_drop_down.get_attribute('value')==program:
                            pass
                        else:
                            xpath = '//div[contains(@id,"ProgramDropDown_DropDown")]//li[contains(text(),"{}")]'.format(program)
                            program_to_select = driver.find_element_by_xpath(xpath)
                            program_drop_down = driver.find_element_by_xpath('//input[contains(@id,"ProgramDropDown_Input")]')
                            ActionChains(driver).move_to_element(program_drop_down).click().pause(1).move_to_element(program_to_select).click().perform()
                            stale_flag =0
                            while stale_flag==0:
                                try:
                                    wait.until(EC.staleness_of(program_drop_down))  
                                    stale_flag=1
                                except TimeoutException as ex:
                                    continue
                        #e-invoice and pdf request pages are different, require different checks
                        #to make sure that the invoice is available
                        #so, in the following if-else I check first that the invoice is available then I get it
                        if page == invoice_request_page:
                            date_checker = driver.find_element_by_xpath('//span[contains(@id,"AvailableQuarterLabelValue")]')
                            if date_checker.text == yq2:
                                cont_flag = 0
                                codes = driver.find_elements_by_xpath('//li[contains(@id,"_ELabelerCodeListBox_")]')
                                ActionChains(driver).move_to_element(codes[0]).click().key_down(Keys.SHIFT).move_to_element(codes[-1]).click().key_up(Keys.SHIFT).perform()
                                submit = driver.find_elements_by_xpath('//input[@value="Submit" and @type="submit" and contains(@id,"InvoiceSubmitButton_input")]')[1]
                                retrieved.append(state+' '+program+' '+'CMS Format')
                            else:
                                cont_flag = 1
                        else:
                            print('Parsing page')
                            soup2 = BeautifulSoup(driver.page_source,'html.parser')
                            print('Obtaining dates')
                            dates = [x.text for x in soup2.find_all('li') if len(x.text)==len(yq3)]
                            #makes sure the drop down has the dates
                            print('checking dates')
                            if any(yq3 in x for x in dates):
                                print('Current quarter available')
                                cont_flag = 0
                                year_quarter_select = driver.find_element_by_xpath('//input[contains(@id,"PIFYearQuarterDropDown_Input")]')                
                                year_quart_to_select = driver.find_element_by_xpath('//div[contains(@id,"PIFYearQuarterDropDown_DropDown")]//li[text()="{}"]'.format(yq3))                   
                                ActionChains(driver).move_to_element(year_quarter_select).click().pause(1).move_to_element(year_quart_to_select).click().pause(1).perform()
                                codes = driver.find_elements_by_xpath('//div[@title="Select LabelerCode"]//li[contains(@id,"_ctl00_LabelerCodeListBox_")]')
                                print('Codes Selected, attempting download')
                                ActionChains(driver).move_to_element(codes[0]).click().key_down(Keys.SHIFT).move_to_element(codes[-1]).click().key_up(Keys.SHIFT).perform()
                                retrieved.append(state+' '+program+' '+'PDF Format')
                                submit = driver.find_elements_by_xpath('//input[@value="Submit" and @type="submit" and contains(@id,"InvoiceSubmitButton_input")]')[1]
                            else:
                                print('date check failed')
                                cont_flag = 1
                        if cont_flag == 1:
                            continue
                        else:
                            pass             
                        submit.click()
                        stale_flag =0
                        while stale_flag==0:
                            try:
                                wait.until(EC.staleness_of(submit))
                                stale_flag=1
                            except TimeoutException as ex:
                                continue
            except Exception as ex:
                print(ex.with_traceback)
                i_accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_RadCheckBoxAccept"]/span[1]')))
                i_accept.click()
                user_name = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtUserName"]')
                user_name.send_keys(username)
                pass_word = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtPassword"]')
                pass_word.send_keys(password)
                login = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_btnLogin_input"]')
                login.click()    
                invoice_request_page().click()
        driver.close()         
        os.chdir('O:\\')
        os.removedirs(self.temp_folder_path)
        return retrieved, to_address

                
def main():
    grabber = PrimsGrabloid()
    requested,to_address = grabber.pull()
    
if __name__=='__main__':
    main()
    
