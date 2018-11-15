# -*- coding: utf-8 -*-
"""
Created on Fri Nov  9 11:05:11 2018

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
import multiprocessing as mp
from multiprocessing.pool import Pool
from grabloid import Grabloid
from concurrent.futures import ProcessPoolExecutor
from IllinoisGrabloid import IllinoisGrabloid
@push_note
def main():
    grabber = IllinoisGrabloid()
    grabber.download_reports()
    grabber.driver.close()
    grabber.cleanup()
if __name__ == "__main__":
    main()