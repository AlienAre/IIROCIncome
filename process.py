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

	#--------- Handle cycle income --------
	incomefile = "F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\Income" + enddate.strftime('%m%d%Y') +".txt"
	print incomefile
	#---- read file to df --------
	if os.path.isfile(incomefile):
		rowno = 0
		with open(incomefile) as f:
				for line in f: 
					rowno = rowno + 1
					if re.match(enddate.strftime('%Y%m%d'), line.lstrip(' ')):
						dataline = re.match(enddate.strftime('%Y%m%d'), line.lstrip(' ')).group()
						print dataline
						print 'data starts at row ' + str(rowno)
						break
						
		dfincome = pd.read_csv(incomefile, sep='|', engine='python', header=None, skiprows=rowno-1, skipfooter=1)			
		dfincome.dropna(axis=1, how='all', inplace=True)
		print dfincome.head()
		print dfincome.shape
	else:
		print 'The cycle income file you need is not saved yet, please save the file first'

	dfhead = pd.read_excel("F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Reference\\Header.xlsx") #get cycle income column header
	columns = dfhead['Header'].tolist()
	dfincome.columns = columns
	print dfincome.head()
		
	#--------- Handle RO advance --------	
	advancefile = "F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\Advance " + enddate.strftime('%m%d%Y') +".csv"
	print advancefile
	#---- read file to df --------
	if os.path.isfile(advancefile):
		rowno = 0
		with open(advancefile) as f:
				for line in f: 
					rowno = rowno + 1
					if re.match(enddate.strftime('%Y%m%d'), line.lstrip(' ')):
						dataline = re.match(enddate.strftime('%Y%m%d'), line.lstrip(' ')).group()
						print dataline
						print 'data starts at row ' + str(rowno)
						break
						
		dfincome = pd.read_csv(incomefile, sep='|', engine='python', header=None, skiprows=rowno-2, skipfooter=1)			
		dfincome.dropna(axis=1, how='all', inplace=True)
		print dfincome.head()
		print dfincome.shape
	else:
		print 'The cycle income file you need is not saved yet, please save the file first'

	#--------- Handle Garnishee -------------
	garnisheefile = "F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\Advance " + enddate.strftime('%m%d%Y') + ".csv"
	print garnisheefile
	if os.path.isfile(garnisheefile):
		rowno = 0
		with open(garnisheefile) as f:
			for line in f:
				rowno = rowno + 1
				if re.match(enddate.strftime('%'))
	
	
	#--------- db connection ----------------
	odbc_conn_str = "DSN=DSDPRD;DBQ=DSDPRD;UID=wwang;PWD=west33" #connect to DSDB using your usr and pssd
	conn = pyodbc.connect(odbc_conn_str) #open co

	#--------- get cycle cslts info 
	sql = '''
	SELECT DISTINCT 
		BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_NUM
		,BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_STATUS
		,BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_SDLR_NUM
		,BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_TERM_DTE
	FROM BRANUSER.BRAN_LKG_CSLT
	WHERE BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_AREA_NUM <> ? 
	AND BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_SDLR_NUM = ?
	AND BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_SMPL_DTE = ? 
	'''	
	
	print 'read sql'
	df = pd.read_sql(sql, conn, params=['6','9737',enddate.strftime("%d-%b-%Y")]) #Oracel accept '15-Aug-2018' format
	print df.head()
