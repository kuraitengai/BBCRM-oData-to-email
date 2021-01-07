# -*- coding: utf-8 -*-
"""
Created on Mon Oct 12 10:49:34 2020

@author: test
"""

import requests
import pandas as pd
import datetime
from secret import username
from secret import password

def assignmenterrors():
    SERVICE_URL = 'ODATA_URL'
    
    fullusername = 'S26\\' + username
    
    session = requests.Session()
    session.auth = (fullusername, password)
    
    session.get(SERVICE_URL)
    
    r = session.get(SERVICE_URL)
    
    r_json = r.json()
    
    df = pd.json_normalize(r_json['value'])
    #print(df)
    
    if r_json['odata.count'] == '0':
        empty = {'': ["There are no errors."], ' ': ["Keep up the good work."]}
        df = pd.DataFrame(data=empty)
    else:
        df = pd.json_normalize(r_json['value'])
        df = df.drop(['QUERYRECID'], axis=1)
        df = df[['LookupID','Name','ProspectManager','StartDate','SpouseManager','SpouseStartDate']]
        df = df.sort_values(['ProspectManager','StartDate'])
        df['StartDate'] = pd.to_datetime(df['StartDate'])
        df['SpouseStartDate'] = pd.to_datetime(df['SpouseStartDate'])
        df['StartDate'] = df['StartDate'].dt.date
        df['SpouseStartDate'] = df['SpouseStartDate'].dt.date
    
    return df

