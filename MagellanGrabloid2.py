# -*- coding: utf-8 -*-
"""
Created on Tue Jul 17 08:28:59 2018

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

class MagellanGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Magellan')

    
    def pull(self):
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        login_credentials = self.credentials
        states = {
            'AK': 'Alaska',
            'AL': 'Alabama',
            'AR': 'Arkansas',
            'AS': 'American Samoa',
            'AZ': 'Arizona',
            'CA': 'California',
            'CO': 'Colorado',
            'CT': 'Connecticut',
            'DC': 'District of Columbia',
            'DE': 'Delaware',
            'FL': 'Florida',
            'GA': 'Georgia',
            'GU': 'Guam',
            'HI': 'Hawaii',
            'IA': 'Iowa',
            'ID': 'Idaho',
            'IL': 'Illinois',
            'IN': 'Indiana',
            'KS': 'Kansas',
            'KY': 'Kentucky',
            'LA': 'Louisiana',
            'MA': 'Massachusetts',
            'MD': 'Maryland',
            'ME': 'Maine',
            'MI': 'Michigan',
            'MN': 'Minnesota',
            'MO': 'Missouri',
            'MP': 'Northern Mariana Islands',
            'MS': 'Mississippi',
            'MT': 'Montana',
            'NA': 'National',
            'NC': 'North Carolina',
            'ND': 'North Dakota',
            'NE': 'Nebraska',
            'NH': 'New Hampshire',
            'NJ': 'New Jersey',
            'NM': 'New Mexico',
            'NV': 'Nevada',
            'NY': 'New York',
            'OH': 'Ohio',
            'OK': 'Oklahoma',
            'OR': 'Oregon',
            'PA': 'Pennsylvania',
            'PR': 'Puerto Rico',
            'RI': 'Rhode Island',
            'SC': 'South Carolina',
            'SD': 'South Dakota',
            'TN': 'Tennessee',
            'TX': 'Texas',
            'UT': 'Utah',
            'VA': 'Virginia',
            'VI': 'Virgin Islands',
            'VT': 'Vermont',
            'WA': 'Washington',
            'WI': 'Wisconsin',
            'WV': 'West Virginia',
            'WY': 'Wyoming',
            'Absolute' : 'South Carolina',
            'BlueChoice' :'South Carolina',
            'First' :'South Carolina',
            'Unison' :'Ohio'
        }
        #make sure the directory is the downloads folder!
        
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        to_address = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Notification Address', usecols='A',dtype='str',names=['Email'],header=None).iloc[0,0]
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Magellan', usecols='D,F',dtype='str')
        mapper = dict(zip(mapper['CLD Programs'],mapper['Lilly Code']))
        #Login with provided credentials
        driver.get('https://mmaverify.magellanmedicaid.com/cas/login?service=https%3A%2F%2Feinvoice.magellanmedicaid.com%2Frebate%2Fj_spring_cas_security_check')   
        user_name = driver.find_element_by_xpath('//*[@id="username"]')
        user_name.send_keys(username)
        pass_word = driver.find_element_by_xpath('//*[@id="password"]')
        pass_word.send_keys(password)
        wait2 = WebDriverWait(driver,3)
        login_button = driver.find_element_by_xpath('//*[@id="content"]/div/div[2]/fieldset/ol[2]/li/input[3]')
        login_button.click()
        '''
        Now moving onto invoices
        '''
        
        #These lines of code get all available options
        invoices_tab = wait.until(EC.element_to_be_clickable((By.ID,'mainForm:invoices')))
        invoices_tab.click()
        business_line = lambda: driver.find_element_by_id('mainForm:srchBusinessLine')
        business_line_select = lambda: Select(business_line())
        business_line_types = [x.text for x in business_line_select().options]
        program_name_options = []
        program_name = lambda: driver.find_element_by_id('mainForm:srchProgramName')
        program_name_select = lambda: Select(program_name())
        for biz in business_line_types:
            business_line_select().select_by_visible_text(biz)
            time.sleep(2)
            _ = [x.text for x in program_name_select().options]
            for x in _:
                program_name_options.append(x)
        year_qtr = lambda: driver.find_element_by_id('mainForm:srchYearQtr')

        
        issue_list = []
        retrieved = {k:[] for k in dict.fromkeys(program_name_options)}
        cld_to_get = []
        already_have = []
        invoices_obtained = []
        '''
        for root, dirs, files in os.walk(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\Invoices'):
            already_have.append(root)
        
        already_have = [x.split('\\')[-3] for x in already_have if len(x.split('\\'))>9 and x.split('\\')[-1]=='Q'+str(qtr)]
        '''
        #Now starting to loop through the options and downloading the files
        #start of business line loop
        for biz in business_line_types:
            print("Working on "+biz+" files")
            wait.until(EC.element_to_be_clickable((By.ID,'mainForm:srchProgramName')))
            business_line_select().select_by_visible_text(biz)
            time.sleep(2)
            program_name = lambda: driver.find_element_by_id('mainForm:srchProgramName')
            program_name_select = lambda: Select(program_name())
            program_name_options = [x.text for x in program_name_select().options]
            time.sleep(1)
            year_qtr().clear()
            year_qtr().send_keys(str(yr)+str(qtr))
            time.sleep(1)
            #start of program name loop. Logic to skip if
            #there are no invoices located here
            for program in program_name_options:
                if program in already_have:
                    print('Invoice already obtained for '+program)
                    continue
                else:
                    pass
                print('Working on '+program+' invoices.')
                program_name_select().select_by_visible_text(program)
                search = driver.find_element_by_xpath('//*[@id="srchInvoiceDiv"]/ol[2]/li/input')
                search.click()
                time.sleep(1)
                invoices = lambda: driver.find_elements_by_xpath('//table//tbody//tr//td//input')
                names = [x.get_attribute('name') for x in invoices()]
                if len(names)==0:
                    continue
                else:
                    pass
                invoice_labels = lambda: driver.find_elements_by_xpath('//table//tbody//tr//td[string-length()>0][1]')
                invoice_labels = [x.text for x in invoice_labels()]
                ps = program.split(' ')[0]
                if ps =='New':
                    if 'York' in ps:
                        state = 'NY'
                    else:
                        state = 'NH'
                elif len(program.split(' ')[0]) <3:
                    state = program.split(' ')[0]
                elif program.split(' ')[0] in ('BlueChoice','First'):
                    state = 'SC'
                elif program.split(' ')[0]=='Unison':
                    state = 'OH'
                elif program.split(' ')[0]=='North':
                    state = 'NC'
                elif program.split(' ')[0] =='Arkansas':
                    state = 'AR'
                else:
                    state = program.split(' ')[0]
                try:
                    lilly_code = mapper[program]
                except KeyError as err:
                    lilly_code = program
                #Loop through the available invoices for the program
                for inv_name, label in zip(names, invoice_labels):
                    invoice = lambda: driver.find_element_by_name(inv_name)
                    invoice().click()
                    labeler_code = label.split('-')[1]
                    cld_info = (labeler_code,program)
                    cld_to_get.append(cld_info)
                    print('Downloading '+label)
                    time.sleep(1)
                    invoice_options = lambda: driver.find_element_by_id('mainForm:selectedFormatType')
                    invoice_options_select = lambda: Select(invoice_options())
                    invoice_options_options = [x.text for x in invoice_options_select().options]
                    #Get both the PDF and the CMS file for the invoice
                    for i,option in enumerate([invoice_options_options[0],invoice_options_options[-1]]):
                        invoice_options_select().select_by_visible_text(option)                 
                        continue_button = wait.until(EC.element_to_be_clickable((By.ID,'mainForm:continueButton')))
                        continue_button.click()
                        time.sleep(5)
                        if i ==0:
                            if 'Invoice Report .pdf' in os.listdir():
                                success_flag = 1
                            else: 
                                pass
                            print('Downloading PDF format.')
                            try:
                                zzz = wait2.until(EC.element_to_be_clickable((By.XPATH,'//a[@href="mailto:rebate@magellanhealth.com"]')))
                                issue_text = program+' '+label+' PDF was not downloaded due to website error, looping unitl downloaded'
                                print(issue_text)
                                issue = [program,label]
                                print('a')
                                driver.back()
                                success_flag = 0
                                count = 0
                                print('b')
                                while (success_flag ==0 and count <10):
                                    if 'Invoice Report .pdf' in os.listdir():
                                        success_flag = 1
                                    else:
                                        pass
                                    driver.refresh()
                                    print('c')
                                    wait.until(EC.element_to_be_clickable((By.NAME,inv_name)))
                                    invoice().click()
                                    invoice_options_select().select_by_visible_text(option)
                                    if 'Invoice Report .pdf' in os.listdir():
                                        success_flag = 1
                                    else:
                                        pass
                                    print('d')
                                    continue_button = driver.find_element_by_id('mainForm:continueButton')
                                    continue_button.click()
                                    print('e')
                                    time.sleep(5)
                                    kount=0
                                    while ('Invoice Report .pdf' not in os.listdir() and kount <10):
                                        print('f')
                                        time.sleep(1)
                                        kount+=1  
                                    count +=1
                                    print('g')
                                    try:
                                        print('h')
                                        zzz = wait2.until(EC.element_to_be_clickable((By.XPATH,'//a[@href="mailto:rebate@magellanhealth.com"]')))
                                        driver.back()
                                        try:
                                            print('i')
                                            wait2.until(EC.visibility_of_element_located((By.ID,'suggestions-list')))
                                            driver.refresh()
                                        except TimeoutException as ec:
                                            pass
                                    except TimeoutException as ex:
                                        print('j')
                                        if "Invoice Report .pdf" in os.listdir():
                                            success_flag = 1
                                        else:
                                            pass
                                if count >10:
                                    print('Tried to get PDF invoice for ' + program+' moving onto next')
                                    issues_list.append(issue)
                                else:
                                    print('Download success after '+str(count)+' tries!')
                                    pass                                       
                            except TimeoutException as ex:
                                pass
                            else:
                                pass
                            path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Invoices\\'+states[state]+'\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                            try:
                                lilly_code = mapper[program]
                            except KeyError as err:
                                lilly_code = program
                            file_name = f'{state}_{lilly_code}_{qtr}Q{yr}_{label.split("-")[1]}.pdf'
                            if os.path.exists('path')==False:
                                os.makedirs(path, exist_ok=True)
                            else:
                                pass
                            if file_name in os.listdir(path):
                                file_name = f'{state}_{lilly_code}_{qtr}Q{yr}_{label.split("-")[1]}_{len(os.listdir(path))+1}.pdf'
                            shutil.move("Invoice Report .pdf",path+file_name)
                            invoices_obtained.append(file_name)
                            retrieved[program].append(label)
                            time.sleep(1)                   
                        else:
                            print('Downloading CMS format.')
                            try:
                                zzz = wait2.until(EC.element_to_be_clickable((By.XPATH,'//a[@href="mailto:rebate@magellanhealth.com"]')))
                                issue = program+' '+label+' CMS was not downloaded due to website error, please try again later.'
                                print(issue)
                                driver.get('https://einvoice.magellanmedicaid.com/rebate/spring/main?execution=e2s1')
                                invoices_tab = driver.find_element_by_id('mainForm:invoices')
                                invoices_tab.click()
                                year_qtr().clear()
                                year_qtr().send_keys(str(yr)+str(qtr))
                                business_line_select().select_by_visible_text(biz)
                                program_name_select().select_by_visible_text(program)
                                search = driver.find_element_by_xpath('//*[@id="srchInvoiceDiv"]/ol[2]/li/input')
                                search.click()
                                time.sleep(2)
                                invoice().click()
                                issue_list.append(issue)
                                time.sleep(5)
                                continue
                            except TimeoutException as ex:
                                pass
                            while 'einvoice.txt' not in os.listdir():
                                time.sleep(1)
                            path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Invoices\\'+states[state]+'\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                            file_name = f'{state}_{lilly_code}_{qtr}Q{yr}_{label.split("-")[1]}.txt'
                            if os.path.exists('path')==False:
                                os.makedirs(path, exist_ok=True)
                            else:
                                pass
                            if file_name in os.listdir(path):
                                file_name = f'{state}_{lilly_code}_{qtr}Q{yr}_{label.split("-")[1]}_{len(os.listdir(path))+1}.txt'
                            shutil.move("einvoice.txt",path+file_name)
                            invoices_obtained.append(file_name)
                            retrieved[program].append(label)
                            time.sleep(1)
                    invoice().click()
        return cld_to_get, invoices_obtained
         ########################################################CLD Below this line###########################################   
        
    def pull_cld(self,cld_to_get):
        states = {
            'AK': 'Alaska',
            'AL': 'Alabama',
            'AR': 'Arkansas',
            'AS': 'American Samoa',
            'AZ': 'Arizona',
            'CA': 'California',
            'CO': 'Colorado',
            'CT': 'Connecticut',
            'DC': 'District of Columbia',
            'DE': 'Delaware',
            'FL': 'Florida',
            'GA': 'Georgia',
            'GU': 'Guam',
            'HI': 'Hawaii',
            'IA': 'Iowa',
            'ID': 'Idaho',
            'IL': 'Illinois',
            'IN': 'Indiana',
            'KS': 'Kansas',
            'KY': 'Kentucky',
            'LA': 'Louisiana',
            'MA': 'Massachusetts',
            'MD': 'Maryland',
            'ME': 'Maine',
            'MI': 'Michigan',
            'MN': 'Minnesota',
            'MO': 'Missouri',
            'MP': 'Northern Mariana Islands',
            'MS': 'Mississippi',
            'MT': 'Montana',
            'NA': 'National',
            'NC': 'North Carolina',
            'ND': 'North Dakota',
            'NE': 'Nebraska',
            'NH': 'New Hampshire',
            'NJ': 'New Jersey',
            'NM': 'New Mexico',
            'NV': 'Nevada',
            'NY': 'New York',
            'OH': 'Ohio',
            'OK': 'Oklahoma',
            'OR': 'Oregon',
            'PA': 'Pennsylvania',
            'PR': 'Puerto Rico',
            'RI': 'Rhode Island',
            'SC': 'South Carolina',
            'SD': 'South Dakota',
            'TN': 'Tennessee',
            'TX': 'Texas',
            'UT': 'Utah',
            'VA': 'Virginia',
            'VI': 'Virgin Islands',
            'VT': 'Vermont',
            'WA': 'Washington',
            'WI': 'Wisconsin',
            'WV': 'West Virginia',
            'WY': 'Wyoming',
            'Absolute' : 'South Carolina',
            'BlueChoice' :'South Carolina',
            'First' :'South Carolina',
            'Unison' :'Ohio'
        }
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        claims_details = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainForm:claims"]')))
        claims_details.click()
        yq = str(yr)+str(qtr)
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Magellan', usecols='D,F',dtype='str')
        mapper = dict(zip(mapper['CLD Programs'],mapper['Lilly Code']))
            
        """
        Sets dropdown default to null
        """
        labeler = lambda: driver.find_element_by_id('mainForm:labelerCode')
        labeler_select = lambda: Select(labeler())
        year_qtr = lambda: driver.find_element_by_id('mainForm:srchYearQtr')
        year_qtr().clear()
        year_qtr().send_keys(yq)
        program_name = lambda: driver.find_element_by_id('mainForm:srchProgramName')
        program_name_select = lambda: Select(program_name()) 
        wait2 = WebDriverWait(driver,3)
        for item in cld_to_get:
            labeler_code = item[0]
            cld_program_name = item[1]
            year_qtr().clear()
            year_qtr().send_keys(yq)
            #input the labeler code and program name from the invoices
            email_flag = 0
            labeler_select().select_by_visible_text(labeler_code)
            time.sleep(1)
            program_name_select().select_by_visible_text(cld_program_name)
            time.sleep(1)
            #VA requires a checkbox, so use a try clause to move past it
            try: 
                a = driver.find_element_by_xpath('//input[@type="checkbox"]')
                a.click()
            except NoSuchElementException as ex:
                pass
            print(item[1]+' selected')
            #click the submit button
            submit = wait.until(EC.element_to_be_clickable((By.ID,'mainForm:btnContinue')))
            submit.click()
            time.sleep(3)
            print('submit clicked')
            #sometimes the site wants to email you when the data is ready, 
            #so switch to that notificaiton and accept if required
            try:
                alert = driver.switch_to.alert
                alert.accept()
                email_flag=1
            except:       
                pass
            #If for some reason the CLD doesn't exist detect the error message
            #add the CLD to the issues list to be sent to the user and move on
            try:
                driver.find_element_by_class_name('errorMsg')
                print('No data for this program')
                driver.refresh()
                continue
            except NoSuchElementException as ex:
                pass
            if email_flag ==0:
                success_flag = 0
                while success_flag ==0:
                    try:
                        wait2.until(EC.element_to_be_clickable((By.XPATH,'//p[contains(text(),"We apologize for the inconvenience and appreciate your patience.")]')))
                        driver.back()
                        submit = wait.until(EC.element_to_be_clickable((By.ID,'mainForm:btnContinue')))
                        submit.click()
                        wait.until(EC.staleness_of((submit)))
                    except TimeoutException as ex:
                        success_flag=1
                        pass
                while 'claimdetails.xls' not in os.listdir():
                    time.sleep(1)
                if len(item[1].split(' ')[0]) <3:
                    state = item[1].split(' ')[0]
                elif item[1].split(' ')[0] in ('BlueChoice','First','Absolute'):
                    state = 'SC'
                elif item[1].split(' ')[0]=='Unison':
                    state = 'OH'
                elif item[1].split(' ')[0]=='North':
                    state = 'NC'
                elif 'New Hampshire' in item[1]:
                    state = 'NH'
                elif 'New York' in item[1]:
                    state = 'NY'
                elif 'Arkansas' in item[1]:
                    state = 'AR'
                else:
                    state = item[1].split(' ')[0]
                    
                path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\'+'Claims\\'+states[state]+'\\'+' '.join(item[1].split(' ')[1:])+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                if os.path.exists(path)==False:
                    os.makedirs(path)              
                else:
                    pass
                lilly_code = mapper[cld_program_name]
                new_name = f'{state}_{lilly_code}_{qtr}Q{yr}_{labeler_code}.xls'
                shutil.move('claimdetails.xls',path+new_name)
            else:
                pass
        os.chdir('O:\\')
        os.removedirs(self.temp_folder_path)
        return invoices

def main():
    grabber = MagellanGrabloid()
    cld, invoices = grabber.pull()
    grabber.pull_cld()
    grabber.send_message(invoices)
    
if __name__=='__main__':
    main()



