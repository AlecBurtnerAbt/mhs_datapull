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
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        login_credentials = self.credentials
        password = login_credentials.iloc[0,1]
        username = login_credentials.iloc[0,0]
        
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
        state_dropdown_func = Select(state_dropdown_func())

        
        # Bypass first value because it is not a state
        states = [x.text for x in state_dropdown.options][1:]                
        for state in states:
            state_dropdown_func.select_by_visible_text(state)
            type_select = Select(driver.find_element_by_xpath('//select[@id="document_document_type_id"]'))
            invoice_types = [x.text for x in type_select.options][1:]
                #Skip the first invoice type because it is not a report 
                



































a = UtahGrabloid()
driver = a.driver
