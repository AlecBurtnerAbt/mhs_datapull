# -*- coding: utf-8 -*-
"""
Created on Mon Jul 16 10:08:13 2018

@author: C252059
"""
# Load all applicable modules
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

class FlexGrabloid(Grabloid):
    def __init__(self):
        super().__init__(script='DrugRebate.com')
        
        
    def pull(self):
        yr = self.yr
        qtr = self.qtr
        driver = self.driver
        wait = self.wait
        mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Delaware', usecols=[0,1],dtype='str')
        login_credentials = self.credentials
        password = self.password
        username = self.username
        states = {
        'AC': 'Maryland',
        'AK': 'Alaska',
        'AL': 'Alabama',
        'AR': 'Arkansas',
        'AS': 'American Samoa',
        'AZ': 'Arizona',
        'BC': 'Maryland',
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
        'MC': 'Maryland',
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
    
        driver.get('https://www.drugrebate.com/RebateWeb/login.do')
        driver.maximize_window()
        user_name = driver.find_element_by_xpath('//*[@id="username"]')
        user_name.send_keys(username)
        pass_word = driver.find_element_by_xpath('//*[@id="password"]')
        pass_word.send_keys(password)
        login_button= driver.find_element_by_xpath('//*[@id="loginBtn"]')
        login_button.click()
        ##Inside page, got to reports
        
        menu = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="SearchInvoices"]/span/a')))
        menu.click()
        
        
        #Leave labeler code blank and it will pull all of them
        quarter = wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id = "startQq"]')))
        quarter.send_keys(str(qtr))    
        year = driver.find_element_by_xpath('//input[@id = "qtrYear"]')    
        year.send_keys(str(yr))
        
        submit_button = driver.find_element_by_xpath('//input[@value="Submit"]')
        submit_button.click()    
        
        #The above block of code got all of the invoics for all labels.
        #Now we have to create the reference dictionary to name files after they
        #are downloaded and then download the files.
        
        soup = BeautifulSoup(driver.page_source,'html.parser')
        data = soup.find_all('td')
        columns = soup.find_all('th')
        columns = [x.text for x in columns]
        data = [x.text for x in data]
        data = np.asarray(data)
        data = data.reshape(-1,9)
        data_asframe = pd.DataFrame(data,columns=columns)
        all_invoices = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@name="downloadWhat"][@value="all"]')))
        all_invoices.click()    
        download_button = driver.find_element_by_xpath('//input[@value="Download Invoices"]')
        download_button.click()
        
        try:
            alert = driver.switch_to.alert
            alert.accept()
        except NoSuchElementException as ex:
            pass
        while any(map((lambda x: 'invoice' in x),os.listdir()))==False:
            time.sleep(1)
        invoices_obtained = []
                    
            
        
        '''
        The below snippet goes to the downloads and unzips all the downloads from the above loop.
        After unzipping the file the zip file gets deleted. 
        '''
        flag = 0
        while flag==0:
            file = os.listdir()[0]
            if file[-3:] != 'zip':
                pass
            else:
                flag=1
        zips = os.listdir()
        zips = [file for file in zips if 'zip' in file]
        
        for file in zips:
            flag = 0
            while flag ==0:
                try:
                    with zipfile.ZipFile(file,'r') as zip_ref:
                        zip_ref.extractall()
                    os.remove(file)
                    flag=1
                except PermissionError as ex:
                    pass
        #now that the files have been unzipped they will be renamed,
        # and moved to the LAN folder
        files = os.listdir()
        for file in files:
            key = file.split('.')[0]
            abbrev = file.split('.')[1][:2]
            state = states[abbrev]
            file_type_abbrev = file.split('.')[2]
            program = file.split('.')[1][2:]
            labeler = list(data_asframe['Labeler'][data_asframe['Invoice']==key])[0]
            if file_type_abbrev =='NCO':
                file_type = 'Claims'
            else:
                file_type = 'Invoices'
            ext = file[-4:]
            path = f'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\{file_type}\\{state}\\{program}\\{yr}\\Q{qtr}\\'
            file_name = f'{abbrev}_{program}_{qtr}Q{yr}_{labeler}{ext}'
            if os.path.exists(path)==False:
                os.makedirs(path)
            else:
                pass
            shutil.move(file,path+file_name)
            if file_type=='Invoices':
                invoices_obtained.append(file_name)
        driver.close()
        os.chdir('O:\\')
        os.removedirs(self.temp_folder_path)
        return invoices_obtained
    
    def morph_cld(self):
        paths = ['O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\Montana\\',
                 'O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\New Mexico\\']
        data = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Test\Raw Text\Montana New Mexico-Conduent Text CLD\DRAMS NCPDP Format.xls',skiprows=2)
        column_names = data.iloc[:,0]
        data_cuts = [(int(start)-1,int(end)-1) for start,end in zip(data.Start,data.End)]
        for path in paths:
            os.chdir(path)
            for root, folders, files in os.walk(path):
                for file in files:
                    try:
                        data = []
                        with open(root+'\\'+file) as F:
                            lines = F.readlines()[1:]
                            for line in lines:
                                holder = []
                                for start,end in data_cuts:
                                    holder.append(line[start:end])
                                data.append(holder)
                        mexontana_frame = pd.DataFrame(data, columns=column_names)
                        new_name = root+'\\'+file[:-3]+'xlsx'
                        mexontana_frame.to_excel(new_name,index=False)
                        os.remove(root+'\\'+file)
                    except UnicodeDecodeError as err:
                        data = []
                        with open(root+'\\'+file,encoding='latin 1') as F:
                            lines = F.readlines()[1:]
                            for line in lines:
                                holder = []
                                for start,end in data_cuts:
                                    holder.append(line[start:end])
                                data.append(holder)
                        mexontana_frame = pd.DataFrame(data, columns=column_names)
                        file_name = file.split('.')[0]+'.xlsx'
                        new_name = root+'\\'+file_name
                        mexontana_frame.to_excel(new_name,index=False)
                        os.remove(root+'\\'+file)
def main():
    grabber = FlexGrabloid()
    invoices = grabber.pull()
    grabber.morph_cld()
    grabber.send_message(invoices)
    
if __name__=='__main__':
    main()
    
    
    
    
    
    
    