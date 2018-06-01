import os, re, sys, time, xlrd, pyodbc, datetime
from datetime import date
import fnmatch
import numpy as np
import pandas as pd
import itertools as it
from openpyxl import load_workbook
from shutil import copyfile
import myfun as dd

def get_tbldate(driver, db_file, sql):
	odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
	conn = pyodbc.connect(odbc_conn_str)
	#--------------------------------------------------------------------
	#Cdatedf = pd.read_sql_query(sql,conn)
	#latestcycledate = Cdatedf.at[(0, 'CDate')]
	cursor = conn.cursor()
	latestcycledate = cursor.execute(sql).fetchone().LDate
	cursor.close()
	conn.close()
	return latestcycledate

def update_tbldate(driver, db_file, sql):
	odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
	conn = pyodbc.connect(odbc_conn_str)
	cursor = conn.cursor()
	cursor.execute(sql)
	cursor.commit()
	cursor.close()
	conn.close()	
	
def add_to_tbl(driver, db_file, tbl, cols, df):
	odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

	for row in df.to_records(index=False):
		values = ", ".join(['\'%s\'' % x for x in row])
		values = values.replace("'nan'", "NULL")
		
		#print values
		sql = '''INSERT INTO %s (%s) VALUES (%s);'''
		sql = sql % (tbl, cols, values)
		#print sql 
		#sys.exit("done")
		conn = pyodbc.connect(odbc_conn_str)
		cursor = conn.cursor()
		cursor.execute(sql)
		cursor.commit()
		cursor.close()
	conn.close()	
	
def try_float(num):
    try:
        floatnum = float(num)
    except ValueError:
        return num
    else:
        return floatnum

def split_str_to_col (str):
	#strip left blanks and end '\n', '\r'
	ltstr = filter(None, re.split(r'\s{2,}', str.lstrip(' ').rstrip('\n').rstrip('\r')))

	for idx in range(len(ltstr)):
		#remove front and trailing blank for each element and remove ',' seperator for numbers
		ltstr[idx] = ltstr[idx].lstrip(' ').rstrip(' ').replace(',', '')
		#update '-' to front to show correct negitive amount
		if '-' in ltstr[idx]:
			try:
				float(ltstr[idx].replace('-', ''))
			except ValueError:
				ltstr[idx]
			else:
				ltstr[idx] = float('-' + ltstr[idx].replace('-', ''))
		ltstr[idx] = try_float(ltstr[idx])	
	#ltstr = [try_float(x) for x in ltstr]		
	return ltstr

def transfer_txt_to_ds (str):	
	dfoutputdata = []	
	cycledate = ''
	outputnamedate = ''
	outputname = '' #use for accumulator num
	
	with open(str) as f:
		for line in f:
			if line.strip():
				#print 'in strip'
				# get cycle end date
				if CycDatePa.match(line.lstrip(' ')) and len(cycledate) == 0:
					#get start date and end date to a list, set cycledate to end date
					#cycledate = re.search('20\d{2}\s+\D{3}\s+\d{1,2}', line).group()
					cycledate = re.findall('20\d{2}\s+\D{3}\s+\d{1,2}', line)[1]
					#print cycledate
					tempd = datetime.datetime.strptime(cycledate, '%Y %b %d')
					cycledate = tempd.strftime('%m/%d/%Y')
					outputnamedate = tempd.strftime('%Y%m%d')
					#break
				# get ACCUMULATOR TYPE 
				if accumtyppattern.match(line[2:].lstrip(' ')) and len(outputname) == 0:
					outputname = re.search(r'\d+', line[2:]).group()			
					#break
				#get normal data lines	
				if DataPa.match(line[2:].lstrip(' ')):
					dfoutputdata.append(split_str_to_col(line[2:]))
				#get totals from file
				if re.match(r'TOTAL ACCUMULATED AMOUNT', line[2:].lstrip(' ')):
					filetotal = split_str_to_col(line[2:])
					#print filetotal
						
	#print 'before assign'
	labels = ['CNSLT NUM', 'CACT TYPE', 'CURRENT DEALERSHIP', 'IGFS ACCUMULATED AMOUNT', 'IGSI ACCUMULATED AMOUNT', 'TOTAL ACCUMULATED AMOUNT']
	df = pd.DataFrame(dfoutputdata, columns=labels)
		
	df['CYCLE END DATE'] = cycledate#datetime.datetime.strptime(cycledate, '%m/%d/%Y')

	print 'now handle ' + outputname
	#print df.dtypes
	if np.isclose(df['IGSI ACCUMULATED AMOUNT'].sum(), float(filetotal[2])):
		print 'IGFI ACCUMULATED AMOUNT matches' 
	if np.isclose(df['TOTAL ACCUMULATED AMOUNT'].sum(), filetotal[3]):
		print 'TOTAL ACCUMULATED AMOUNT matches' 	

	return df

	
	
#------ program starting point --------	
if __name__=="__main__":
	## dd/mm/yyyy format
	print 'Process date is ' + str(time.strftime("%m/%d/%Y"))
	print 'Please enter the cycle end date (mm/dd/yyyy) you want to process:'
	#-----------------------------------------------------
	#------- get cycle date ----------------------
	getcycledate = datetime.datetime.strptime(raw_input(), '%m/%d/%Y')
	endday = getcycledate
	startday = dd.getCStartDate(getcycledate)

	print 'Cycle start date is ' + str(startday)
	print 'Cycle end date is ' + str(endday)

	#--------- database info ----------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
	db_file = r"F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\IIROC_Income.accdb;"
	#db_file = r"C:\\pycode\\IIROCIncome\\IIROC_Income.accdb;"
	user = "admin"
	password = ""
	#--------------------------------------------------------------------
	#--------- check latest cycle date in database ----------
	sql = '''SELECT Max([CDate]) AS [LDate] FROM [tCycleDate]; '''
	latestcycledate = get_tbldate(driver, db_file, sql)
	print 'In database, the latest cycle end date is ' + str(latestcycledate)
		
	if latestcycledate >= endday:
		print 'It seems that the cycle date in database is later than your cycle date, which means the transactions may already be entered into database. Please type "1" if you do not want to stop and check:'
		if raw_input() == '1':
			sys.exit("The process is stopped")
	#-----------------------------------------------------------------------
	
	#--------- update cycle date table to latest cycle end date ----------
	sql = '''Update [tCycleDate] SET [CDate] = #''' + str(endday.strftime("%m/%d/%Y")) + '''#; '''
	update_tbldate(driver, db_file, sql)
	#-----------------------------------------------------------------------
	
	copyfile('F:\\Files For\\Hai Yen Nguyen\\IIROC reporting\\IIROC_Income.accdb', 'M:\\bak\\IIROC_Income ' + str(date.today().strftime("%m%d%Y")) + '.accdb')
	#copyfile('C:\\pycode\\IIROCIncome\\IIROC_Income.accdb', 'C:\\pycode\\IIROCIncome\\IIROC_Income' + str(date.today().strftime("%m%d%Y")) + '.accdb')
	
	#DataPa = re.compile(r'^\d{1,5}\s{2,}\d{1}\s{2,}IG\D{2}\s{1}\(\d{4}\).*$')
	DataPa = re.compile(r'\d{1,5}\s{2,}\d{1}\s{2,}\D{4}.*$')
	accumtyppattern = re.compile(r'ACCUMULATOR TYPE.*\d+')
	CycDatePa = re.compile(r'.*THRU\s+20\d{2}\s+\D{3}\s+\d{1,2}')
	Negative = re.compile(r'-')

	debtlist = ['1', '2', '7', '1193']
	iiroclist = ['33', '35', '794']
	allist = ['335', '602', '696', '1133', '1177', '1181', '1317', '1334', '1336', '1338', '1340']

	filelist = [] 

	for file in os.listdir('C:\\pycode\\IIROCIncome'):
		if file.endswith('.txt'):
			filelist.append(os.path.join('C:\\pycode\\IIROCIncome', file))

	#print filelist	
	for txts in filelist:
		with open(txts) as f:
			for line in f:
				if line.strip():
					if accumtyppattern.match(line[2:].lstrip(' ')):
						accumulatortype = re.search(r'\d+', line[2:]).group()
						print 'Will start to process  ACCUMULATOR TYPE ' + accumulatortype
						if accumulatortype == '7':
							dfoutput = transfer_txt_to_ds(txts)
							
							sql = '''SELECT Max([CycleDate]) AS [LDate] FROM [tAdvances_Historical]; '''						
							tbldate = get_tbldate(driver, db_file, sql)
							print 'The lastest cycle date is ' + tbldate.strftime('%m/%d/%Y')
							if tbldate >= endday:
								print 'It seems that the data for accumulatortype ' + accumulatortype + ' you are inserting has been in database. Please type "1" to stop and check:'
								if raw_input() == '1':
									sys.exit("The process is stopped")
							table = '''tAdvances_Historical'''
							columns = '''[CSLT_Num], [CACT], [Dealership], [IGFS_Amt], [IGSI_Amt], [Total_Amt], [CycleDate]'''
							add_to_tbl(driver, db_file, table, columns, dfoutput)
						elif accumulatortype == '11':
							dfoutput = transfer_txt_to_ds(txts)
							
							sql = '''SELECT Max([CycleDate]) AS [LDate] FROM [tNetPay_Historical]; '''						
							tbldate = get_tbldate(driver, db_file, sql)
							print 'The lastest cycle date is ' + tbldate.strftime('%m/%d/%Y')
							if tbldate >= endday:
								print 'It seems that the data for accumulatortype ' + accumulatortype + ' you are inserting has been in database. Please type "1" to stop and check:'
								if raw_input() == '1':
									sys.exit("The process is stopped")
							table = '''tNetPay_Historical'''
							columns = '''[CSLT_Num], [CACT], [Dealership], [IGFS_Amt], [IGSI_Amt], [Total_Amt], [CycleDate]'''
							add_to_tbl(driver, db_file, table, columns, dfoutput)								
						elif accumulatortype == '151':
							dfoutput = transfer_txt_to_ds(txts)
							
							sql = '''SELECT Max([CycleDate]) AS [LDate] FROM [tGarnishedBalance_Historical]; '''						
							tbldate = get_tbldate(driver, db_file, sql)
							print 'The lastest cycle date is ' + tbldate.strftime('%m/%d/%Y')
							if tbldate >= endday:
								print 'It seems that the data for accumulatortype ' + accumulatortype + ' you are inserting has been in database. Please type "1" to stop and check:'
								if raw_input() == '1':
									sys.exit("The process is stopped")
							table = '''tGarnishedBalance_Historical'''
							columns = '''[CSLT_Num], [CACT], [Dealership], [IGFS_Amt], [IGSI_Amt], [Total_Amt], [CycleDate]'''
							add_to_tbl(driver, db_file, table, columns, dfoutput)		

						break
	print 'The import process is completed succefully'					
					