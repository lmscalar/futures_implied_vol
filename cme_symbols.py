# -*- coding: utf-8 -*-
"""
Created on Mon May 23 11:03:12 2016

@author: lmolina
"""

import pandas as pd
import numpy as np
import xlwings as xw


df = pd.read_excel(r'C:\python_modules\curves\cmePrices\NG.xlsm','Symbols', header=0,index_col=1, parse_cols= 'B:C')
cols = [ s.replace(" ","_").lower() for s in df.columns]
df.columns = cols

F = pd.HDFStore(r'C:\risk_data\mc_scenarios.h5','a')
 
F['cme_symbols'] = df

F.close()

K = pd.HDFStore(r'C:\risk_data\mc_scenarios.h5','r')

dg = K['cme_symbols']

K.close()