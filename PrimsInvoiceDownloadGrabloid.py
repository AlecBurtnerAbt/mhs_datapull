# -*- coding: utf-8 -*-
"""
Created on Tue Aug 21 09:46:33 2018

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
    This grabloid navigates the Prims site by Molina health care.
    The page is a total pain in the ass because many of the elements
    in it are rendered by Ajax.  The class method pull(self) not only downloads
    the invoices but also parses them for NDCs, associates those
    NDCs to programs, and then goes and requests those CLD files
    '''
    def __init__(self):
        super().__init__(script='Prims')

    def pull(self):
        statesII = {
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
        username = login_credentials.iloc[0,0]
        password = login_credentials.iloc[0,1]
        driver = self.driver
        driver.get('https://primsconnect.dxc.com')
        wait = self.wait
        i_accept = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_RadCheckBoxAccept"]/span[1]')))
        i_accept.click()
        flex_mapper = pd.read_excel(r'O:\M-R\MEDICAID_OPERATIONS\Electronic Payment Documentation\Automation Scripts Parameters\automation_parameters.xlsx',sheet_name='Prims', usecols='D,E,F,G',dtype='str',names=['state','flex_id','state_id','state_name'])
        user_name = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtUserName"]')
        user_name.send_keys(username)
        pass_word = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_txtPassword"]')
        pass_word.send_keys(password)
        login = driver.find_element_by_xpath('//*[@id="ctl00_PlaceHolderMain_LoginWebPart_ctl00_btnLogin_input"]')
        login.click()          
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
            state_selected_flag = 0
            while state_selected_flag == 0:
                print('Checking state drop down value')
                drop_down = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@title="Select State"]')))
                if drop_down.get_attribute('value')==state:
                    print('State already selected')
                    pass
                else:
                    print('Selecting state...')
                    drop_down = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@title="Select State"]')))
                    ActionChains(driver).move_to_element(drop_down).click().pause(5).perform()
                    xpath = '//div[contains(@id,"ctl00_StateDropDown_DropDown")]//li[text()="{}"]'.format(state)
                    state_to_select = driver.find_element_by_xpath(xpath)
                    ActionChains(driver).move_to_element(state_to_select).click().pause(5).perform()
                    wait.until(EC.staleness_of(drop_down))
                    print('State selected.')
                drop_down = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@title="Select State"]')))
                if drop_down.get_attribute('value')==state:
                    print('State selected, passing value to state selected flag')
                    state_selected_flag = 1
                soup = BeautifulSoup(driver.page_source,'html.parser')
                print('Parsing lists...')
                lists2 = soup.find_all('ul',attrs={'class':'rcbList'})
                programs = [x.text.split('-') for x in lists2[1]]
                codes = [x[0].strip() for x in programs]
                name = ['-'.join(x[1:]) for x in programs]
                programs = dict(zip(codes,name))
                print('Program mapper updated')
                state_programs.update({state:programs})        
        driver.back() 
        try:
            alert = driver.switch_to.alert
            alert.accept()
        except NoAlertPresentException as ex:
            pass
        
        #The below block of code crawls through available download pages and 
        #downloads the data, renames it, and moves it to the appropriate directory
       
        
        #Chane the number of reports per page
        number_per_page = wait.until(EC.presence_of_element_located((By.XPATH,'//input[contains(@id,"ChangePageSizeTextBox")]')))
        number_per_page.clear()    
        number_per_page.send_keys('9999')    
        inter=driver.find_element_by_xpath('//input[@type="submit" and @value="Change"]')
        inter.click()
        #Get the pages downloads are on
        pages = lambda: driver.find_elements_by_xpath('//div[@class="rgWrap rgNumPart"]/a')     
        for i,page in enumerate(pages()):
            p = pages()[i]
            pages()[i].click()
            soup = BeautifulSoup(driver.page_source,'html.parser')
            table = soup.find('table',attrs={'class':'rgMasterTable'})
            body = table.find_all('tbody')[1]            
            data = [x.text.strip() for x in body.find_all('td')]
            data = np.asarray(data)
            data = data.reshape(-1,9)
            data = data[data[:,-1]=='Download']
            data = pd.DataFrame(data,columns=['report_id','manufacturer','state','date_requested','type','status','file_name','date_complete','download_link'])
            data['state_id'] = data['file_name'].str.rsplit('_',0).str[-1]
            data['state_id'] = data['state_id'].str.rsplit('.').str[0]
            data= pd.merge(data,flex_mapper,how='left',on=['state','state_id'])
            data = data.fillna('no_flex_id')
            data['labeler'] = data['file_name'].str.split('_').str[2]
            data_copy = data
            data = data[(data['type']=='Invoice')|(data['type']=='EInvoice')]
            data = data.drop_duplicates(subset=['file_name']).reset_index(drop=True)
            
            ndc_list = []
            for i in range(len(data)):
                state = data.loc[i,'state']
                program = data.loc[i,'flex_id']
                file_type = data.loc[i,'file_name'][-4:].lower()
                labeler= data.loc[i,'labeler']
                if program =='no_flex_id':
                    program = state_programs[statesII[state]][data.loc[i,'state_id']].strip()
                path ='O:\\M-R\\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Invoices\\{}\\{}\\{}\\Q{}\\'.format(statesII[state],program,yr,qtr)
                file_name = '_'.join([state,program,yq2,labeler])+file_type
                if os.path.exists(path)==False:
                    os.makedirs(path)
                xpath = '//tr/td[text()="{}"]/following-sibling::td/span[contains(@id,"_lnkDownload")]'.format(data.loc[i,'file_name'])
                link = driver.find_element_by_xpath(xpath)
                link.click()
                while data.loc[i,'file_name'] not in os.listdir():
                    time.sleep(1)
                if file_type =='.txt':
                    read_flag =0
                    while read_flag ==0:
                        try:
                            with open(data.loc[i,'file_name']) as f:
                                lines = f.readlines()
                                menu_item = '  -  '.join([data.loc[i,'state_id'],state_programs[statesII[state]][data.loc[i,'state_id']].strip()])
                                ndcs = (state,menu_item,list(set([x[6:17] for x in lines])))
                                ndc_list.append(ndcs)
                                read_flag=1
                        except PermissionError as ex:
                            pass
                shutil.move(data.loc[i,'file_name'],path+file_name)
                invoices_obtained.append(file_name)
                
       
        #Have the state : NDC tuples, move onto getting CLD
        submit_request = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@value="Submit Request"]')))
        submit_request.click()     
        
        
        
        
        for report in ndc_list:
            print(f'State is {report[0]}')
            state_drop_down = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@title="Select State"]')))
            if state_drop_down.get_attribute('value')==statesII[report[0]]:
                print('State already selected')
                pass
            else:
                print('Selecting state')
                xpath = '//div[contains(@id,"ctl00_StateDropDown_DropDown") and @class="RadComboBoxDropDown RadComboBoxDropDown_Vista "]//li[text()="{}"]'.format(statesII[report[0]])
                cur_val = state_drop_down.get_attribute('value')
                state_to_select = driver.find_element_by_xpath(xpath)
                ActionChains(driver).move_to_element(state_drop_down).click().pause(2).move_to_element(state_to_select).click().perform()
                wait.until(EC.staleness_of(state_drop_down))
                state_drop_down = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@title="Select State"]')))
                time.sleep(1)
                state_drop_down = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@title="Select State"]')))
                cur_val = state_drop_down.get_attribute('value')
                print(f'wait a tic, state is {statesII[report[0]]} and current value is {cur_val} ')
            selected_flag = 0
    
            soup = BeautifulSoup(driver.page_source,'html.parser')        
            items_labels = [x.text.split(' ')[0] for x in soup.find_all('li') if '-' in x.text and len(x.text)>len(yq3)]
            items_links = driver.find_elements_by_xpath('//li[contains(text(),"-")]')
            letters_to_programs = dict(zip(items_labels,items_links))
            report_code_letter = report[1].split(' ')[0]
            program_to_select = letters_to_programs[report_code_letter]
            program_drop_down = lambda: driver.find_element_by_xpath('//div[contains(@id,"ProgramDropDown")]//input[@title = "Select Program"]')
            ActionChains(driver).move_to_element(program_drop_down()).click().pause(1).move_to_element(program_to_select).click().perform()
            dates_acquired = 0
            try:
                wait.until(EC.staleness_of((program_to_select)))
            except:
                pass
            from_q = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[contains(@name,"$ctl00$FYearQuarterDropDown")]')))
            from_q.click()
            time.sleep(1)
            while dates_acquired==0:
                dates =[x.text for x in driver.find_elements_by_xpath('//div[@id="ctl00_SPWebPartManager1_g_72775f67_6a04_4f06_9334_41276e78dec2_ctl00_FYearQuarterDropDown_DropDown"]//li[contains(text(),"Q")]')]
                if len(dates[0])==0:
                    pass
                else:
                    dates_acquired=1
            if any(yq3 in x for x in dates)==False:
                continue
                #pass
            else:
                from_q = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$FYearQuarterDropDown")]')
                to_q = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$TYearQuarterDropDown")]')
                current_qtr = driver.find_element_by_xpath('//div[@id="ctl00_SPWebPartManager1_g_72775f67_6a04_4f06_9334_41276e78dec2_ctl00_FYearQuarterDropDown_DropDown"]//li[text()="{}"]'.format(yq3))
                current_qtr2 = driver.find_element_by_xpath('//div[@id="ctl00_SPWebPartManager1_g_72775f67_6a04_4f06_9334_41276e78dec2_ctl00_TYearQuarterDropDown_DropDown"]//li[text()="{}"]'.format(yq3))
                ActionChains(driver).move_to_element(from_q).click().pause(1).click(current_qtr).move_to_element(to_q).click().pause(1).move_to_element(current_qtr2).click().perform()
                
            ndcs = ','.join(report[2])
            ndc_box = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$NDCTextBox")]')
            ndc_box.send_keys(ndcs)
            submit = driver.find_element_by_xpath('//input[@type="submit"][@value="Submit"]')
            submit.click()
            wait_flag=0
            while wait_flag==0:
                try:
                    wait.until(EC.staleness_of(submit))
                    wait_flag=1
                except:
                    continue
            ndc_box = driver.find_element_by_xpath('//input[contains(@name,"$ctl00$NDCTextBox")]')
            ndc_box.clear()
        driver.close()
        return invoices_obtained


def main():
    grabber = PrimsDownloadGrabloid()
    invoices = grabber.pull()
    grabber.send_message(invoices)
    grabber.cleanup()
    #send_message(subject=,body=,to='burtner_abt_alec@lilly.com')
    
if __name__=='__main__':
    main()










