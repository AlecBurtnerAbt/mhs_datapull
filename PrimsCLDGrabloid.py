# -*- coding: utf-8 -*-
"""
Created on Thu Oct 25 11:02:21 2018

@author: c252059
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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoAlertPresentException
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

class PrimsDownloadGrabloid(Grabloid):
    '''
    This grabloid class goes to the Prims website by Molina and 
    downloads all the CLD requested by the previously run Prims grabloids
    and then splits it up by program and saves it to the appropriate directory
    '''
    def __init__(self):
        super().__init__(script='Prims')

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
            'WY': 'Wyoming'
                            }
            #Open the webdriver, define the wait function, and get through the login page
        yr = self.yr
        qtr = self.qtr
        login_credentials = self.credentials
        driver = self.driver
        driver.implicitly_wait(30)
        wait = WebDriverWait(driver,30)
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
    
        driver.get('https://primsconnect.dxc.com ')
        fl_flex_mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Prims', usecols='F,G',dtype='str',names=['flex_id','state_id'])
        wv_flex_mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Prims', usecols='Q,R',dtype='str',names=['flex_id','state_id'])
        fl_flex_mapper = dict(zip(fl_flex_mapper['state_id'],fl_flex_mapper['flex_id']))
        wv_flex_mapper = dict(zip(wv_flex_mapper['state_id'],wv_flex_mapper['flex_id']))
        def login_proc():
            i_accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_RadCheckBoxAccept"]/span[1]')))
            i_accept.click()
            user_name = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtUserName"]')
            user_name.send_keys(username)
            pass_word = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtPassword"]')
            pass_word.send_keys(password)
            login = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_btnLogin_input"]')
            login.click()   
        login_proc()
        #Now inside the webpage, begin selection process
        submit_request = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Submit Request"]')))
        submit_request.click()    
        
        yq2 = '{}Q{}'.format(qtr,yr)
        yq3 = '{}-Q{}'.format(yr,qtr)
        #Make the program to state dictionaries
        soup = BeautifulSoup(driver.page_source,'html.parser')
        lists = soup.find_all('ul',attrs={'class':'rcbList'})    
        states = [x.text for x in lists[0]]    
        state_programs = {}
        invoices_obtained = []
        
        invoice_request_page =  wait.until(EC.element_to_be_clickable((By.XPATH,'//span[text()="Electronic Invoice (TXT)"]')))
        invoice_request_page.click()  
        wait.until(EC.presence_of_element_located((By.XPATH,'//td/span[contains(@id,"AvailableQuarterLabelValue")]')))
        #have to select the state to get the state programs to populate
        
        #The below block of code is creating the state: programcode:program name dictionary
        #to create the filenames for after download
        
        
        
        for state in states:
            try:
                print('Checking state drop down value')
                drop_down = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@title="Select State"]')))
                if drop_down.get_attribute('value')==state:
                    print('State already selected')
                    pass
                else:
                    print('Selecting state...')
                    xpath = '//div[contains(@id,"ctl00_StateDropDown_DropDown")]//li[text()="{}"]'.format(state)
                    state_to_select = driver.find_element_by_xpath(xpath)
                    ActionChains(driver).move_to_element(drop_down).click().pause(1).move_to_element(state_to_select).click().perform()
                    time.sleep(7)
                    print('State selected.')
                soup = BeautifulSoup(driver.page_source,'html.parser')
                print('Parsing lists...')
                lists2 = soup.find_all('ul',attrs={'class':'rcbList'})
                programs = [x.text.split('-') for x in lists2[1]]
                codes = [x[0].strip() for x in programs]
                name = ['-'.join(x[1:]) for x in programs]
                programs = dict(zip(codes,name))
                print('Program mapper updated')
                state_programs.update({state:programs})       
            except:
                try:
                    login_proc()
                    invoice_request_page =  wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_rtsRequest"]/div/ul/li[2]/a/span/span/span')))
                    invoice_request_page.click()  
                    wait.until(EC.staleness_of((invoice_request_page)))
                except:
                    continue
       
        try:
            alert = driver.switch_to.alert
            alert.accept()
        except NoAlertPresentException as ex:
            pass
        try:            
            driver.back()
        except:
            login_proc()
        
        #The below block of code crawls through available download pages and 
        #downloads the data, renames it, and moves it to the appropriate directory
       
        
        #Chane the number of reports per page
        #make it so there's only one page to scrape
        number_per_page = driver.find_element_by_xpath('//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radGridRequestSummary_ctl00_ctl03_ctl01_ChangePageSizeTextBox"]')
        number_per_page.clear()    
        number_per_page.send_keys('10000')    
        inter=driver.find_element_by_xpath('//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radGridRequestSummary_ctl00_ctl03_ctl01_ChangePageSizeLinkButton"]')
        inter.click()
        
        
        #Build the dataframe of files listed as ready
        soup = BeautifulSoup(driver.page_source,'html.parser')
        table = soup.find('table',attrs={'class':'rgMasterTable'})
        body = table.find_all('tbody')[1]            
        data = [x.text.strip() for x in body.find_all('td')]
        data = np.asarray(data)
        data = data.reshape(-1,9)
        data = data[data[:,-1]=='Download']
        data = pd.DataFrame(data,columns=['report_id','manufacturer','state','date_requested','type','status','file_name','date_complete','download_link'])
        data['state_id'] = data.file_name.str.split('_').str[7]
        data= pd.merge(data,flex_mapper,how='left',on=['state','state_id'])
        data = data.fillna('no_flex_id')
        claim_data = data[(data['type']=='Claims')&(data['status']=="Ready for download")]
        claim_data['labeler'] = claim_data.file_name.str.split('_').str[2].str[:5]
        claim_data = claim_data.drop_duplicates(subset='file_name').reset_index(drop=True)
        
        
        #iterate through the dataframe to download files
        for i in range(len(claim_data)):
            file_name = data.loc[i,'file_name']
            xpath = '//tr/td[text()="{}"]/following-sibling::td/span[contains(@id,"_lnkDownload")]'.format(claim_data.loc[i,'file_name'])
            DL_flag=0
            #loop to insure that inactivity on the web page leading to 
            #logout won't stop the program
            while DL_flag ==0:
                print(f'Start of download process for {file_name}')
                download_link = driver.find_element_by_xpath(xpath)
                download_link.click()
                counter = 0
                while file_name not in os.listdir() and counter <10:
                    time.sleep(3)
                    counter +=1
                #if file in directory, success
                if file_name in os.listdir():
                    print(f'Downloaded {file_name}')
                    DL_flag=1
                else:
                    try:
                        #have to reset teh page to conditions for reclicking on 
                        #the download link
                        print('Logged out, logging in')
                        i_accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_RadCheckBoxAccept"]/span[1]')))
                        i_accept.click()
                        user_name = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtUserName"]')
                        user_name.send_keys(username)
                        pass_word = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtPassword"]')
                        pass_word.send_keys(password)
                        login = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_btnLogin_input"]')
                        login.click() 
                        print('Resetting page conditions')
                        number_per_page = driver.find_element_by_xpath('//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radGridRequestSummary_ctl00_ctl03_ctl01_ChangePageSizeTextBox"]')
                        number_per_page.clear()    
                        number_per_page.send_keys('10000')    
                        inter=driver.find_element_by_xpath('//*[@id="ctl00_SPWebPartManager1_g_967e6faf_f673_482f_95d3_d22fbf4faf7a_ctl00_radGridRequestSummary_ctl00_ctl03_ctl01_ChangePageSizeLinkButton"]')
                        inter.click()
                    except:
                        print('New error, chcek the web page')
                        
        #Due to the logic dealing with logouts go through the directory and make sure 
        #all duplicate files are removed
        for file in os.listdir():
            if '(' and ')' in file:
                print(f'Removing duplicate file {file}')
                os.remove(file)
        
        # Now that duplicate files are gone, read files into dataframes
        # One dataframe for each state because their file layouts are different
        names = list(set(['{}_{}'.format(x.split('_')[1],x.split('_')[7]) for x in os.listdir()]))    
        frames = {k:pd.DataFrame() for k in names}
        files = [x.split('_') for x in os.listdir()]    
        files = sorted(files,key=lambda x: (x[1],x[7]))
        files = ['_'.join(x) for x in files]
        for file in files:
            state = file.split('_')[1]
            program_code = file.split('_')[7]
            key = '{}_{}'.format(state,program_code)
            if state=='FL':
                skip = 1
            else:
                skip = 3
            temp = pd.read_excel(file,skiprows=skip,skipfooter=1)
            frames[key] = frames[key].append(temp)
            #os.remove(file)
      
        for key in frames.keys():
            state = key.split('_')[0]
            program_code = key.split('_')[1]
            if state == 'FL':
                program = fl_flex_mapper[program_code]
            else:
                program = wv_flex_mapper[program_code]
            path ='O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\{}\\{}\\{}\\Q{}\\'.format(statesII[state],program,yr,qtr)
            if os.path.exists(path):
                pass
            else:
                os.makedirs(path)
            file_name = '{}_{}_{}Q{}.xlsx'.format(state,program,qtr,yr)
            os.chdir(path)
            frames[key].to_excel(file_name,engine='xlsxwriter',index=False)
        os.chdir('O:\\')
        os.removedirs(self.temp_folder_path)
def main():
    grabber = PrimsDownloadGrabloid()
    grabber.pull()

if __name__=='__main__':
    main()
'''
driver = grabber.driver        
qtr, yr, login_credentials = grabber.qtr, grabber.yr, grabber.credentials        
wait = grabber.wait        
'''        
        
        
        
        
        
        
        
    
    
    