#--------------------------------------------
#version:		1.0
#author:		West
#Description:	used to prepare IIROC Income Report
#				Report run by each cycle.
#				Data source are cycle income, RO advance, Garnishee balance and Net pay
#				Need all active IIROC cslts and cslts who terminated within 30 days; information can be
#				obtained from DSDB, Branuser table
#Workflow:		read all data files based on the cycle date requested,
#				get all cslts information from DSDB based on the cycle date requestedfilter
#				combine all the data and filter out inactive cslts Data
#				output to excel
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

sys.path.append("C:\\pycode\\libs")
import igtools as ig
import dbquery as dbq

#------ program starting point --------
if __name__=="__main__":
	## dd/mm/yyyy format
	print("This is the process for IIROC Income Reporting")
	print("Process date is " + str(time.strftime("%m/%d/%Y")))
	#print("Please enter the cycle end date (mm/dd/yyyy) you want to process:")
	#-----------------------------------------------------
	#------- get cycle date ----------------------
	#enddate = datetime.datetime.strptime(raw_input(), "%m/%d/%Y")
	enddate = datetime.datetime.strptime(input("Please enter the cycle end date (mm/dd/yyyy) you want to process:"), "%m/%d/%Y")
	startdate = ig.getCStartDate(enddate)

	print("Cycle start date is " + str(startdate))
	print("Cycle end date is " + str(enddate))

	#--------- Handle Garnishee -------------
	garnisheefile = "C:\\pycode\\IIROCIncome\\Garnishee " + enddate.strftime("%m%d%Y") + ".csv"
	print(garnisheefile)

	if os.path.isfile(garnisheefile):
		total = float(0)
		file = csv.reader(open(garnisheefile, newline=""), delimiter=",")

		#get the total amount "Report Totals" from file, used to check the total of all data rows
		for line in file:
			if len(line) > 0:
				if "Report Totals" in line[0]:
					total = ig.str2float(line[3])
					break

		#get all data rows and format
		dfgarnishee = pd.read_csv(garnisheefile, engine="python", skiprows=6, skipfooter=2)
		dfgarnishee[["Cslt No.","CACT"]].astype("int64")
		dfgarnishee["Total Amount"] = dfgarnishee["Total Amount"].apply(ig.str2float)
		dfgarnishee["Cycle End Date"] = dfgarnishee["Cycle End Date"].astype("datetime64")

		#make sure total of all data rows matches the total amount "Report Totals" from file
		if total != dfgarnishee["Total Amount"].sum().round(2):
			print(total, " <> ", dfgarnishee["Total Amount"].sum().round(2))
			print("please check your file")
			sys.exit()
		else:
			print(total, " = ", dfgarnishee["Total Amount"].sum().round(2))
			dfsumgarnishee = dfgarnishee.groupby(["Cslt No."])["Total Amount"].sum().round(2)
	else:
		print("It seems the file you need is not saved yet, please save the file first.")

	#--------- Handle RO advance -------------
	netpayfile = "C:\\pycode\\IIROCIncome\\NetPay " + enddate.strftime("%m%d%Y") + ".csv"

	if os.path.isfile(netpayfile):
		total = float(0)
		file = csv.reader(open(netpayfile, newline=""), delimiter=",")

		#get the total amount "Report Totals" from file, used to check the total of all data rows
		for line in file:
			if len(line) > 0:
				if "Report Totals" in line[0]:
					total = ig.str2float(line[3])
					break

		#get all data rows and format
		dfnetpay = pd.read_csv(netpayfile, engine="python", skiprows=6, skipfooter=2)
		dfnetpay[["Cslt No.","CACT"]].astype("int64")
		dfnetpay["Total Amount"] = dfnetpay["Total Amount"].apply(ig.str2float)
		dfnetpay["Cycle End Date"] = dfnetpay["Cycle End Date"].astype("datetime64")
