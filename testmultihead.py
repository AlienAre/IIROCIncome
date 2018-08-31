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

sys.path.append('C:\\pycode\\libs')
import igtools as ig
import dbquery as dbq

#------ program starting point --------	
if __name__=="__main__":

	#--------- Handle cycle income --------
	incomefile = "F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\Income08152018.txt"
	dfincome = pd.read_csv(incomefile, sep='|', engine='python', header=[0,1], skipfooter=1)			
	print dfincome.head()
	print dfincome.tail()
	
	writer = pd.ExcelWriter("Income.xlsx", engine="xlsxwriter")
	
	dfincome.to_excel(writer, index=False)

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()
	
	print 'The process is done'