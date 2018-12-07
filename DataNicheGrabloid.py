# -*- coding: utf-8 -*-
"""
Created on Wed Dec  5 12:36:28 2018

@author: c252059
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Dec  5 10:42:29 2018

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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementNotVisibleException
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

class DataNicheGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script="DataNiche")
        
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
            'DC': 'Wash DC',
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
            'S0': 'South Carolina'
        }
        driver = self.driver
        driver.get('https://dna-outlierview3.imshealth.com/Login')
        user = self.credentials.iloc[0,0]
        password = self.credentials.iloc[0,1]
        wait = self.wait
        states_to_get = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='DataNiche', usecols='D',dtype='str')
        user_name_input = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@name="username"]')))
        user_name_input.send_keys(user)
        password_input = driver.find_element_by_xpath('//input[@id="password-field"]')
        password_input.send_keys(password)
        login_button = driver.find_element_by_xpath('//input[@value="LOG IN"]')
        login_button.click()
        #Find the button bar and the select button
        select_button = wait.until(EC.element_to_be_clickable((By.XPATH,'//div[contains(@class,"btn-group btn-block")]/button[2]')))
        select_button.click()
        full_states = [states[x] for x in states_to_get.iloc[:,0]]
        '''
        Have to implement while loops to ensure that the bot can navigate
        to the state's tab correctly.  This means a lot of nested whiles.
        This is the start of the state
        '''
        for state in full_states:
            print(f'State is {state}')
            sidebar_link = wait.until(EC.presence_of_element_located((By.XPATH,f'//table[@id="statetbl"]//td[text()="{state}"]')))
            ActionChains(driver).move_to_element(sidebar_link).click().pause(8).perform()
            print(f'Clicked on {state}')
            #Identify programs and begin program loop            
            programs = lambda: driver.find_elements_by_xpath('//div[@id="forReview"]/div[@ng-repeat="program in programs"]/div/div[contains(@class,"myInvoiceB")]/div[9]/button')
            for i,program in enumerate(programs()):
                print(f'Program is {programs()[i].text}')
                ActionChains(driver).move_to_element(programs()[i]).click().perform()
                time.sleep(6)
                wait.until(EC.presence_of_element_located((By.XPATH,'//div[@class="slimScrollDiv"]')))
                labeler_tabs = lambda: driver.find_elements_by_xpath('//div[@class="slimScrollDiv"]//li')[1:]                
                #We now have the labeler tabs, time to loop
                #through the tabs and approve the data                
                for j, labeler in enumerate(labeler_tabs()):
                    print(f'Labeler is {labeler_tabs()[j].text}')
                    ActionChains(driver).move_to_element(labeler_tabs()[j]).click().perform()
                    time.sleep(6)
                    approve_button = driver.find_element_by_xpath("""//button[@ng-click="ApproveOrRejectVerified('approve')"]""")
                    approve_button.click()
                    time.sleep(8)
                    if j != 2:
                        ActionChains(driver).move_to_element(programs()[i]).click().perform()
                        time.sleep(6)
                    else:
                        pass
                    


                
    
            
            for program in programs:
                program.click()


def main():
    grabber = DataNicheGrabloid()
    grabber.pull()



if __name__ == "__main__":
    main()





<span class="type9 mLeft30 ng-binding">Connecticut</span>


<button class="col-sm-3 btn btn-lg approve-btn type5 greenB approvebtnalign toggleArrow threehundwidth-btn" ng-click="ApproveOrRejectVerified('approve')">
                    <span>
                        <img src="../images/approve_green.png">
                        <img src="../images/approve_white.png">
                    </span>&nbsp;APPROVE
                </button>




