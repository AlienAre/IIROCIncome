#--------------------------------------------
#version:		1.0.0.1
#author:		West
#Description:	used to prepare annual consultant benefit credit
#Assumptions:	report end date is 03/31 each year
#				New Business is past 12 trailing months sales credits/new business 
# 
#--------------------------------------------

import os, re, sys, csv, time, xlrd, pyodbc, datetime
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
	## dd/mm/yyyy format
	print 'This is the process for IIROC Income Reporting'
	print 'Process date is ' + str(time.strftime("%m/%d/%Y"))
	print 'Please enter the cycle end date (mm/dd/yyyy) you want to process:'
	#-----------------------------------------------------
	#------- get cycle date ----------------------
	enddate = datetime.datetime.strptime(raw_input(), '%m/%d/%Y')
	startdate = ig.getCStartDate(enddate)
	
	print 'Cycle start date is ' + str(startdate)
	print 'Cycle end date is ' + str(enddate)
	
	#--------- Handle Garnishee -------------
	garnisheefile = "F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\Garnishee " + enddate.strftime('%m%d%Y') + ".csv"
	print garnisheefile
	
	if os.path.isfile(garnisheefile):
		file = csv.reader(open(garnisheefile, 'rb'), delimiter=',')
		for line in file:
			print line
#			if re.search(r'Report Totals', line):
#				summary = line.split(',')
#				print summary
#
	