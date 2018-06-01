#--------------------------------------------
#version:		1.0.0.1
#author:		West
#Description:	used to prepare annual consultant benefit credit
#Assumptions:	report end date is 03/31 each year
#				New Business is past 12 trailing months sales credits/new business 
# 
#--------------------------------------------

import os, re, sys, time, xlrd, pyodbc, datetime
from datetime import date
import fnmatch
import numpy as np
import pandas as pd
import itertools as it
from openpyxl import load_workbook
from shutil import copyfile

df = pd.read_csv('C:\\Users\\wangwe5\\Documents\\Download\\Income05152018.txt', sep='|', header=None, skiprows=2)

print df.head()
df.drop([95], axis=1, inplace=True)
print df.head()
#writer = pd.ExcelWriter('Income.xlsx', engine='xlsxwriter')
#df.to_excel(writer, index=False)
#writer.save()

