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
	
	incomefile = "F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\Income08152018.txt"
	rowno = 0
	with open(incomefile) as f:
		for line in f: 
			rowno = rowno + 1
			if re.match("20180815", line.lstrip(" ")):
				a = re.match("20180815", line.lstrip(" ")).group()
				print a
				print rowno
				break	
	df = pd.read_csv("F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\Download\\Income08152018.txt", sep="|", engine="python", header=None, skiprows=rowno-1, skipfooter=1)			
	df.drop([95], axis=1, inplace=True)

	print df.head()
	print df.tail()
	print df.shape
	
	writer = pd.ExcelWriter("Income.xlsx", engine="xlsxwriter")
	
	df.to_excel(writer, index=False)

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()
	
	print 'The process is done'
