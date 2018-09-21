#--------------------------------------------
#version:		1.0
#author:		West
#Description:	used to prepare IIROC Income Report
#				Report run by each cycle.
#				Data source are cycle income, RO advance, Garnishee balance and Net pay
#				Need all active IIROC cslts and cslts who terminated within 30 days; information can be
#				obtained from DSDB, Branuser table
#Workflow:		read all data files based on the cycle date requested,
#				get all cslts information from DSDB based on the cycle date requested filter
#				combine all the data and filter out inactive cslts Data
#				output to excel
#
#--------------------------------------------


import os, re, sys, time, xlrd, pyodbc, datetime
from datetime import date, timedelta
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
	print("This is the process for IIROC Income Reporting")
	print("Process date is " + str(time.strftime("%m/%d/%Y")))
	#-----------------------------------------------------
	#------- get cycle date ----------------------
	enddate = datetime.datetime.strptime(input("Please enter the cycle end date (mm/dd/yyyy) you want to process:"), "%m/%d/%Y")
	startdate = ig.getCStartDate(enddate)

	print("Cycle start date is " + str(startdate))
	print("Cycle end date is " + str(enddate))

	#--------- db connection setting ----------------
	odbc_conn_str = "DSN=DSDPRD;DBQ=DSDPRD;UID=wwang;PWD=west33" #connect to DSDB using your usr and pssd
	conn = pyodbc.connect(odbc_conn_str) #open co

	#--------- sql parameter setting -------------
	area = "6" #Head Office is special, need to be excluded
	dealer = "9737" #IIROC cslts only
	cycdate = enddate.strftime("%d-%b-%Y") #cycle end date; Oracel accept '15-Aug-2018' format
	termdate = enddate - timedelta(days=30)
	print("Termdate is ", termdate)	
	
	#--------- Handle cycle income --------
	incomefile = "F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\Income" + enddate.strftime('%m%d%Y') +".txt"

	#---- read file to df --------
	if os.path.isfile(incomefile):
		rowno = 0
		with open(incomefile) as f:
				for line in f: 
					rowno = rowno + 1
					if re.match(enddate.strftime('%Y%m%d'), line.lstrip(' ')):
						dataline = re.match(enddate.strftime('%Y%m%d'), line.lstrip(' ')).group()
						print(dataline)
						print('data starts at row ' + str(rowno))
						break
						
		dfincome = pd.read_csv(incomefile, sep='|', engine='python', header=None, skiprows=rowno-1, skipfooter=1)			
		dfincome.dropna(axis=1, how='all', inplace=True)
		#print(dfincome.head())
		print(dfincome.shape)
	else:
		print("The cycle income file you need is not saved yet, please save the file first")
		sys.exit()

	dfhead = pd.read_excel("F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Reference\\Header.xlsx") #get cycle income column header
	dfincome.columns = dfhead['Header'].tolist() #assign header to data frame
	dfincome.drop(["CYCLE DATE", "ACCOUNT TYPE", "REP STAT", "APPOINT DATE", "SALES START", "TERMINATE DATE", "AREA NUM", "RO NUM", "DO NUM", "CNSLT AL PAID", "CNSLT INS AL PAID"], axis=1, inplace=True) #remove not required columns
	#dfsumincome = dfincome.sum() #do not work, due to all types are numeric
	dfsumincome = dfincome.groupby("REP NUM").sum().reset_index()
		
	#--------- Handle RO advance --------	
	advancefile = "F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\Advance " + enddate.strftime('%m%d%Y') +".csv"
	dfadvance = ig.read_accumulatorupdated(advancefile)
	dfsumadvance = pd.DataFrame({"Advance": dfadvance.groupby(["Cslt No."])["Total Amount"].sum().round(2)}).reset_index()
	dfsumadvance.columns = ["Cslt", "Advance"]
	
	#--------- Handle Net Pay -------------
	netpayfile = "F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\NetPay " + enddate.strftime('%m%d%Y') + ".csv"
	dfnetpay = ig.read_accumulatorupdated(netpayfile)
	dfsumnetpay = pd.DataFrame({"NetPay": dfnetpay.groupby(["Cslt No."])["Total Amount"].sum().round(2)}).reset_index()
	dfsumnetpay.columns = ["Cslt", "NetPay"]

	
	#--------- Handle Garnishee -------------
	garnisheefile = "F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\Garnishee " + enddate.strftime('%m%d%Y') + ".csv"
	dfgarnishee = ig.read_accumulatorupdated(garnisheefile)
	dfsumgarnishee = pd.DataFrame({"Garnishee": dfgarnishee.groupby(["Cslt No."])["Total Amount"].sum().round(2)}).reset_index()
	dfsumgarnishee.columns = ["Cslt", "Garnishee"]
	
	#--------- get cycle cslts info ---------------
	sql = '''
	SELECT DISTINCT 
		BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_NUM
		,BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_NAM_FULL
		,BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_STATUS
		,BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_SDLR_NUM
		,BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_TERM_DTE
	FROM BRANUSER.BRAN_LKG_CSLT
	WHERE BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_AREA_NUM <> ? 
	AND BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_SDLR_NUM = ?
	AND BRANUSER.BRAN_LKG_CSLT.LKG_CSLT_SMPL_DTE = ? 
	'''	
	
	print("Get cslt info from database")	
	dfcslt = pd.read_sql(sql, conn, params=[area, dealer, cycdate])
	dfoutput = dfcslt[(dfcslt["LKG_CSLT_STATUS"] == "Active") | (dfcslt["LKG_CSLT_TERM_DTE"] > termdate)] #active and cslt terminated within 30 days
	
	#combine all sourse data and filter only required cslts
	print("filter cslts and combine all source data")
	dfoutput.columns = ["Cslt", "Name", "Status", "Dealer", "TermDate"]
	dfoutput = dfoutput.merge(dfsumincome, how="left", left_on="Cslt", right_on="REP NUM")
	dfoutput = dfoutput.merge(dfsumadvance, how="left", on="Cslt")
	dfoutput = dfoutput.merge(dfsumnetpay, how="left", on="Cslt")
	dfoutput = dfoutput.merge(dfsumgarnishee, how="left", on="Cslt")
	print(dfoutput.head())
	
	#output to Excel
	print("prepare output Excel")
	
	writer = pd.ExcelWriter("F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\" + enddate.strftime("%m%d%Y") + " CycleIncome.xlsx", engine="xlsxwriter")
	dfoutput.to_excel(writer, freeze_panes=(1,0), index=False)
	writer.save()
	print("Report has been saved to F:\\Files For\\Hai Yen Nguyen\\IIROC reporting")