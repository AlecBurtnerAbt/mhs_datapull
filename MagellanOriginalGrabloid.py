# -*- coding: utf-8 -*-
"""
Created on Mon Jul 16 14:20:05 2018

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
import re

class MagellanOriginalGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='OldMagellan')
        
    def pull(self):
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        login_credentials = self.credentials
        yq = str(yr)+str(qtr)
        login_credentials = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Magellan', usecols='A,B',dtype='str')
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        to_address = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Notification Address', usecols='A',dtype='str',names=['Email'],header=None).iloc[0,0]
        program_mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='{}'.format(self.script), usecols='D,E',dtype='str')
        program_mapper = dict(zip(program_mapper['Program ID'],program_mapper['Flex Contract ID']))
        wait = WebDriverWait(driver,10)
        driver.get('https://einvoicing.magellanmedicaid.com/rebate')
        user_name = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="input_1"]')))
        user_name.send_keys(username)
        pass_word = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="input_2"]')))
        pass_word.send_keys(password)
        login_button = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="auth_form"]/fieldset/ol[2]/li/input')))
        login_button.click()
        
        '''
        This part of the code grabs the invoices
        '''
        
        links = lambda: driver.find_elements_by_xpath('//table[@id="mainForm:manufacturerTable"]//a')
        url = driver.current_url
        invoices_obtained = []
        for i, link in enumerate(links()):
            links()[i].click()
            invoices = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:invoices"]')))
            invoices.click()
            select_lilly = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="selectedManufacturer"]')))
            select_lilly.click()
            continue_1 = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
            continue_1.click()
            year_quarter = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:selYearDate"]')))
            select = webdriver.support.ui.Select(year_quarter)
            try:
                select.select_by_visible_text(yq)
            except NoSuchElementException as ex:
                continue
            continue_2 = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
            continue_2.click()
            
            #Have to get all of the states available and get their invoices
            wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
            states= driver.find_elements_by_xpath('//input[@id = "selectedClient"]')    
            '''Begin iterating through state options
            '''
            
            for i,state in enumerate(states):
                states[i].click()
                continue_3 = driver.find_element_by_xpath('//*[@id="mainForm:btnContinue"]')
                continue_3.click()
                continue_4 = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
                programs = lambda: driver.find_elements_by_xpath('//input[@type="radio"]')
                '''
                Begin iterating through program options
                '''
                
                for i,program in enumerate(programs()):
                    programs()[i].click()
                    continue_4().click()
                    continue_5 = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
                    codes = lambda: driver.find_elements_by_xpath('//input[@type="checkbox"]')
                    '''
                    Click all available label codes
                    '''
                    [x.click() for x in codes()]
                    continue_5().click()
                    continue_6 = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@type="submit"]')))
                    options = lambda: driver.find_elements_by_xpath('//input[@type="radio" and @class="radioBtn" ]')
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    table = soup.find('ol', attrs={'style':'margin-bottom: 0px;'})
                    lis = table.find_all('li')
                    lis = [x.text.split(':')[1:] for x in lis]
                    lis = [''.join(x).strip() for x in lis]            
                    lis[2] = lis[2].replace('State of','').replace('Commonwealth of','').strip()
                    lis[-1] = lis[-1].replace('\n','_')
                    file_name = '_'.join(lis[1:]).replace('/','-')
                    program = lis[3]
                    
                    if '/' in program:
                        program = program.replace('/','-')
                    '''
                    Getting CMS and PDF output formats and iterating through them
                    '''
                    for i,option in enumerate(options()):
                        path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Invoices\\'+lis[2]+'\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'                           
                        if os.path.exists(path)==False:
                            os.makedirs(path)
                        else:
                            pass
                        if i ==0:
                            option.click()
                            continue_6().click()
                            while 'invoicetext.txt' not in os.listdir():
                                time.sleep(1)                    
                            shutil.move('invoicetext.txt',path+file_name+'.txt')
                        elif i<2:
                            time.sleep(3)
                            option.click()
                            continue_6().click()
                            while 'main.pdf' not in os.listdir():
                                time.sleep(1)
                            shutil.move('main.pdf',path+file_name+'.pdf')
                        elif i >1:
                            pass
                        invoices_obtained.append(file_name)
                    driver.back()
                    driver.back()
                    programs = lambda: driver.find_elements_by_xpath('//input[@type="radio"]')
                driver.back()
                states= driver.find_elements_by_xpath('//input[@id = "selectedClient"]')          
            driver.get(url)
        
        ##############################################The above code gets the invoices
            ################################Below are the claim level details
            
        '''
        Have to make a state reference dictionary for this 
        '''
        state_mapper = {'Commonwealth of Pennsylvania':'PA',
                        'Idaho':'ID',
                        'Nebraska':'NB'}
        links = lambda: driver.find_elements_by_xpath('//table[@id="mainForm:manufacturerTable"]//a')
        url = driver.current_url
        for i, link in enumerate(links()):
            links()[i].click()
            #Goto claims details tab
            claims = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:claims"]')))
            claims().click()
            
            #Select lilly radio button and advance
            select_lilly = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="selectedManufacturer"]')))
            select_lilly().click()
            continue_1 = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
            continue_1.click()
            
            #Enter year and quarter and advance
            year_quarter = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:selYearDate"]')))
            select = lambda: webdriver.support.ui.Select(year_quarter())
            try:
                select().select_by_visible_text(yq)
            except NoSuchElementException:
                print('Claims data not ready for current year and quarter. Terminating program')
                exit()
            continue_2 = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
            continue_2.click()
            
            #Start the state gathering/iterating
            wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
            states= lambda: driver.find_elements_by_xpath('//input[@id = "selectedClient"]') 
            error_programs = []
            for l,state in enumerate(states()):
                time.sleep(1)
                print('a')
                state_number = driver.current_url[-2:]
                state_number = re.sub(r"\D","",state_number)
                states()[l].click()
                continue_3 = lambda: wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:btnContinue"]')))
                continue_3().click()
                programs = lambda: driver.find_elements_by_xpath('//input[@type="radio"]')
                program_codes = [x.text for x in driver.find_elements_by_xpath('//table[@id="mainForm:programTable"]/tbody/tr/td[2]')]
                time.sleep(1)
                print('b')
                
                for j,program,prog_code in zip(range(len(program_codes)),programs(),program_codes):
                    programs()[j].click()
                    program_number = driver.current_url[-2:]
                    program_number = re.sub(r"\D","",program_number)
                    continue_3().click()
                    by_label = lambda: driver.find_element_by_xpath('//*[@id="detailByLabeler"]')
                    by_label().click()
                    print('c')
                    labels = lambda: driver.find_elements_by_xpath('//input[@id="labelerCode"]')
                    soup = BeautifulSoup(driver.page_source,'html.parser')
                    table = soup.find('ol', attrs={'style':'margin-bottom: 0px;'})
                    lis = table.find_all('li')
                    lis = [x.text.split(':')[1:] for x in lis]
                    lis = [''.join(x).strip() for x in lis]            
                    lis[2] = lis[2].replace('State of','').strip()
                    table2 = soup.find('tbody')
                    codes = table2.find_all('td')
                    codes = [x.text.strip() for x in codes if len(x.text)>2]
                    time.sleep(1)
                    print('d')
            
                    for k, label, code in zip(range(0,len(codes)),labels(),codes):
                        
                        success_flag=0
                        labels()[k].click()
                        print('f')
                        file_name = '{}_{}_{}Q{}_{}.xls'.format(state_mapper[lis[2]],program_mapper[prog_code],qtr,yr,code)
                        print('g')
            
                        while success_flag ==0:
                            #added a success_flag because of server side errors found during testing
                            #Sometimes the data isn't available at all, the 'ELIF' statement addresses
                            #that occurence
                            continue_4 = driver.find_element_by_xpath('//input[@id="mainForm:btnContinue2"]')
                            continue_4.click()  
                            time.sleep(2)
                            soup = BeautifulSoup(driver.page_source,'html.parser')
                            error = soup.find('li',attrs={'class':'errorMsg'})
                            state = lis[2]
                            if state=='Commonwealth of Pennsylvania':
                                state = 'Pennsylvania'
                            no_data_error = soup.find()
                            if error == None: 
                                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Claims\\'+state+'\\'+lis[3]+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                                if os.path.exists(path)==False:
                                    os.makedirs(path)
                                else:
                                    pass
                                while 'claimsdetail.xls' not in os.listdir():
                                    time.sleep(1)
                                if ('claimsdetail.xls' in os.listdir())==True:
                                    shutil.move('claimsdetail.xls',path+file_name)
                                    success_flag=1
                                    pass
                                else:
                                    pass
                            elif 'no Claim Level Detail' in error.text:
                                success_flag=1
                                driver.back()
                                continue
                            else:
                                error_programs.append(lis[2]+'_'+lis[3]+'_'+code)
                                by_label().click()
                                time.sleep(2)
                            print('h')
                    driver.get('https://einvoicing.magellanmedicaid.com/rebate/spring/main?execution=e1s%s'%(program_number))
                driver.get('https://einvoicing.magellanmedicaid.com/rebate/spring/main?execution=e1s%s'%(state_number))
            driver.get(url)
        driver.close()
        os.chdir('O:\\')
        os.remove(self.temp_folder_path)
        return invoices_obtained

def main():
    grabber = MagellanOriginalGrabloid()
    invoices = grabber.pull()
    grabber.send_message(invoices)



if __name__=='__main__':
    main()                           

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    