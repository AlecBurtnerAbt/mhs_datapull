# -*- coding: utf-8 -*-
"""
Created on Mon Sep 10 10:35:02 2018

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

class LargeMagellanGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='Magellan')


    def pull(self):
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
        login_credentials = self.credentials
        driver = self.driver
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]        
        #Login with provided credentials
        driver.get('https://mmaverify.magellanmedicaid.com/cas/login?service=https%3A%2F%2Feinvoice.magellanmedicaid.com%2Frebate%2Fj_spring_cas_security_check')   
        user_name = driver.find_element_by_xpath('//*[@id="username"]')
        user_name.send_keys(username)
        pass_word = driver.find_element_by_xpath('//*[@id="password"]')
        pass_word.send_keys(password)
        wait = WebDriverWait(driver,10)
        wait2 = WebDriverWait(driver,3)
        login_button = driver.find_element_by_xpath('//*[@id="content"]/div/div[2]/fieldset/ol[2]/li/input[3]')
        login_button.click()
        
        '''
        Navigate to claims details, requested reports
        '''
        claims_tab = driver.find_element_by_xpath('//a[@id="mainForm:claims"]')
        claims_tab.click()
        
        requested_reports = driver.find_element_by_xpath('//a[@id="mainForm:download"]')
        requested_reports.click()
        
        pages = lambda: driver.find_element_by_xpath('//select[@id="mainForm:reporterPageScroller"]')
        pages_select = lambda: Select(pages())
        page_options = [x.text for x in pages_select().options]
        reports_obtained=[]
        for page in page_options:
            pages_select().select_by_visible_text(page)
            reports =driver.find_elements_by_xpath('//table[@id="mainForm:claimsTable"]//input[@type="submit"]')
            codes = [x.text for x in driver.find_elements_by_xpath('//table[@id="mainForm:claimsTable"]//tr//td[2]')][1:]
            programs =   [x.text for x in driver.find_elements_by_xpath('//table[@id="mainForm:claimsTable"]//tr//td[3]')][1:]      
            states2 = [x.text[:2] for x in driver.find_elements_by_xpath('//table[@id="mainForm:claimsTable"]//tr//td[3]')][1:] 
            for report, code, program, state in zip(reports,codes,programs,states2):
                try:
                    S = states[state]
                    report.click()
                    while 'claimdetails.xls' not in os.listdir():
                        time.sleep(1)
                    file_name = S+'_'+program+'_'+code+'_'+str(qtr)+'Q'+str(yr)+'.xls'
                    path = 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\'+S+'\\'+program+'\\'+str(yr)+'\\'+'Q'+str(qtr)+'\\'
                    if os.path.exists(path)==False:
                        os.makedirs(path)
                    else:
                        pass
                    shutil.move('claimdetails.xls',path+file_name)
                except KeyError as err:
                    pass
                reports_obtained.append(file_name)
        driver.close()
        os.chdir('O:\\')
        os.removedirs(self.temp_folder_path)

def main():
    grabber = LargeMagellanGrabloid()
    grabber.pull()
    
if __name__=='__main__':
    main()








