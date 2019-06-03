# -*- coding: utf-8 -*-
"""
Created on Thu Nov 29 11:23:05 2018

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
from grabloid import Grabloid, push_note
import pickle

class UtahGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script="Utah")
    
    def pull(self, efficient=True):
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
            'Unison' :'Ohio',
            'S0': 'South Carolina',
            'AO':'Arizona'
        }
        states_2 = dict(zip(states.values(),states.keys()))
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        login_credentials = self.credentials
        password = login_credentials.iloc[0,1]
        username = login_credentials.iloc[0,0]
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='{}'.format(self.script), usecols='C,D,E',dtype='str')
        #Navigate to site and log in
        driver.get('https://rsp.ghsinc.com/RebateServicesPortal/application/login.joi')
        username_input = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@name="user_login"]')))
        username_input.send_keys(username)
        password_input = driver.find_element_by_xpath('//input[@name="user_password"]')
        password_input.send_keys(password)
        login_button = driver.find_element_by_xpath('//button[@id="loginFormSubmit"]')
        login_button.click()
        
        
        #Log in takes you directly to invoice page.
        # Have three select boxes, state, type, and period.
        # state must be selected, then type, then period
        
        
        #Build state selection
        state_dropdown = Select(wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="document_state_id"]'))))
        #also assign state_dropdown to a lambda to be able to repeatedly call
        #inside the for loops
        state_dropdown_func = lambda: driver.find_element_by_xpath('//select[@id="document_state_id"]')
        state_dropdown_select = lambda: Select(state_dropdown_func())
        
        
        # Bypass first value because it is not a state
        report_states = [x.text for x in state_dropdown.options][1:]
        state_abbreviations = [states_2[state] for state in report_states]
        state_pairs = [(state,abbrev) for state, abbrev in zip(report_states,state_abbreviations)]
        current_time = f'{yr} - Q{qtr}'                
        for state, abbrev in state_pairs:
            state_dropdown_select().select_by_visible_text(state)
            time.sleep(1)
            type_select = lambda: Select(wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="document_document_type_id"]'))))
            #Skip the first invoice type because it is not a report 
            programs = [x.text for x in type_select().options][1:]
            print(f"Working on {state}'s programs.")
            for program in programs:
                print(f"Working on {program}")
                wait.until(EC.element_to_be_clickable((By.XPATH,f'//select[@id="document_document_type_id"]')))
                time.sleep(1)
                print(program)
                type_select().select_by_visible_text(program)
                time.sleep(1)
                date_select = driver.find_element_by_xpath('//select[@id="document_id"]')
                date_select = Select(date_select)
                available_dates = [x.text for x in date_select.options][1:]                 
                if current_time in available_dates:
                    date_select.select_by_visible_text(current_time)
                else:
                    continue
                submit_button = driver.find_element_by_xpath('//button[@type="submit"]')
                submit_button.click()
                wait.until(EC.staleness_of(submit_button))
                #Now look through the table of results to pick out state, labeler, program
                #also the link to the CMS and PDF format
                rows = driver.find_elements_by_xpath('//tbody[@role="alert"]//tr')
                file_names = [x.text.replace(current_time,f'{qtr}Q{yr}').replace(' ','_') for x in rows]
                cms_links = driver.find_elements_by_xpath('//tbody[@role="alert"]//tr/td[5]/a')
                pdf_links = driver.find_elements_by_xpath('//tbody[@role="alert"]//tr/td[6]/a')
                #Now that we have the links and file names download each file, then move it
                #to the right folder and rename it
                for file_name, cms_link, pdf_link in zip(file_names, cms_links, pdf_links):
                    print(f'Obtaining CMS format for {program}')
                    cms_link.click()
                    while len(os.listdir()) <1:
                        print('Waiting for file to download')
                        time.sleep(1)
                    file = os.listdir()[0]
                    path = f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\{state}\\{program}\\{yr}\\{qtr}'
                    if os.path.exists(path) == False:
                        os.makedirs(path)
                    extension = '.txt'
                    cms_file_name = file_name+extension
                    move_flag = 0
                    while move_flag == 0:
                        try:
                            file = os.listdir()[0]
                            shutil.move(file,os.path.join(path,cms_file_name))
                            move_flag = 1
                        except Exception as ex:
                            print(f'Failed to move {file}')
                            print(ex)
                            time.sleep(1)
                    #have the cms format, now grab pdf
                    print(f'Obtaining PDF format for {program}')
                    pdf_link.click()
                    while len(os.listdir()) <1:
                        print('Waiting for file to download')
                        time.sleep(1)
                    file = os.listdir()[0]
                    extension = '.pdf'
                    pdf_file_name = file_name+extension
                    move_flag = 0
                    while move_flag == 0:
                        try:
                            file = os.listdir()[0]
                            shutil.move(file,os.path.join(path,pdf_file_name))
                            move_flag = 1
                        except Exception as ex:
                            print(f'Failed to move {file}')
                            print(ex)
                            time.sleep(1)

if __name__ =='__main__':
    grabber = UtahGrabloid()
    grabber.pull()
