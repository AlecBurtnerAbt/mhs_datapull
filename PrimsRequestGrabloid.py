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
    
        driver.get('https://www.primsconnect.molinahealthcare.com/_layouts/fba/primslogin.aspx?ReturnUrl=%2f_layouts%2fAuthenticate.aspx%3fSource%3d%252FSitePages%252FHome%252Easpx&Source=%2FSitePages%2FHome%2Easpx')
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
        submit_request = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radLnkSubmitRequest_input"]')))
        submit_request.click()    
        
        #Now in the request page, navigate to invoice tab
        
        invoice_request_page = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_rtsRequest"]/div/ul/li[2]/a/span/span/span')))
        soup = BeautifulSoup(driver.page_source,'html.parser')
        lists = soup.find_all('ul',attrs={'class':'rcbList'})    
        states = [x.text for x in lists[0]]    
        pdf_request_page = lambda: driver.find_element_by_xpath('//span[text()="Paper Invoice (PDF)"]')
        pages = [invoice_request_page,pdf_request_page]
        retrieved = []
        for page in pages:
            try:
                reset_button = driver.find_element_by_xpath('//span[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_EInvoiceResetButton"]')
                page().click()
                wait.until(EC.staleness_of(reset_button))
                
                for state in states:
                    state_input = driver.find_element_by_xpath('//input[contains(@name,"StateDropDown")]')
                    if state_input.get_attribute('value')==state:
                        pass
                    else:
                        state_to_select = driver.find_element_by_xpath('//div[contains(@id,"StateDropDown")]//li[text()="'+state+'"]')
                        ActionChains(driver).move_to_element(state_input).click().pause(1).click(state_to_select).pause(1).perform()
                        wait.until(EC.staleness_of(state_input))
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    lists = soup.find_all('ul',attrs={'class':'rcbList'})
                    programs = [x.text for x in lists[1]]
                    for program in programs:
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
                        if page==invoice_request_page:
                            date_checker = driver.find_element_by_xpath('//span[contains(@id,"AvailableQuarterLabelValue")]')
                            if date_checker.text == yq2:
                                cont_flag = 0
                                codes = driver.find_elements_by_xpath('//li[contains(@id,"_ELabelerCodeListBox_")]')
                                ActionChains(driver).move_to_element(codes[0]).click().key_down(Keys.SHIFT).move_to_element(codes[-1]).click().key_up(Keys.SHIFT).perform()
                                submit = driver.find_element_by_xpath('//span[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_EInvoiceSubmitButton"]')
                                wait.until(EC.element_to_be_clickable((By.XPATH,'//span[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_EInvoiceSubmitButton"]')))
                                retrieved.append(state+' '+program+' '+'CMS Format')
                            else:
                                cont_flag = 1
                        else:
                            print('a')
                            soup2 = BeautifulSoup(driver.page_source,'html.parser')
                            dates = [x.text for x in soup2.find_all('li') if len(x.text)==len(yq3)]
                            if any(yq3 in x for x in dates):
                                print('b')
                                cont_flag = 0
                                year_quarter_select = driver.find_element_by_xpath('//input[contains(@id,"PIFYearQuarterDropDown_Input")]')                
                                year_quart_to_select = driver.find_element_by_xpath('//div[contains(@id,"PIFYearQuarterDropDown_DropDown")]//li[text()="{}"]'.format(yq3))                   
                                ActionChains(driver).move_to_element(year_quarter_select).click().pause(1).move_to_element(year_quart_to_select).click().pause(1).perform()
                                codes = driver.find_elements_by_xpath('//div[@title="Select LabelerCode"]//li[contains(@id,"_ctl00_LabelerCodeListBox_")]')
                                ActionChains(driver).move_to_element(codes[0]).click().key_down(Keys.SHIFT).move_to_element(codes[-1]).click().key_up(Keys.SHIFT).perform()
                                submit = driver.find_element_by_xpath('//span[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_InvoiceSubmitButton"]')
                                wait.until(EC.element_to_be_clickable((By.XPATH,'//span[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_InvoiceSubmitButton"]')))
                                retrieved.append(state+' '+program+' '+'PDF Format')
                            else:
                                print('c')
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
            except:
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
    
