# -*- coding: utf-8 -*-
"""
Created on Tue Oct 16 13:45:52 2018

@author: C252059
"""

import os

path = 'O:\\M-R\MEDICAID_OPERATIONS\\Electronic Payment Documentation\\Test\\Claims\\'

for root, dirs, files in os.walk(path):
    if 'Alabama' not in root:
        pass
    else:
        for file in files:
            print(root+'\\'+file)