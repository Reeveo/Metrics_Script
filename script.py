import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import time
import os


startTime = datetime.now()

def Metrics():
	#Get the raw data
	file = r'FILE PATH'
	if os.path.exists(file) == True:
		print("Found the raw data spreadsheet and it will now be processed and cleaned.")
		raw_data = pd.read_csv(file, header=0, encoding='ISO-8859-1', low_memory= False)
		#Remove any entries which have an error status
		AllData = raw_data[~raw_data['Status'].str.contains('Error', na=False)==True]
		#Select only the columns needed
		AllData = pd.DataFrame(AllData, columns=['COLUMN TITLES])
		print("Columns have been selected.")
		#Select only cases which have specific strings in a specific column
		AllData = AllData[AllData['COLUMN NAME'].str.contains('STRING')==True]
		#Format the date column to date time otherwise it will paste it as text and excel formulas will not work
		AllData['DATE COLUMN NAME'] = pd.to_datetime(AllData['DATE COLUMN NAME'], dayfirst=True)
		# This is the location of the destination file
		XLSXfile = r'FULL PATH OF DESTINATION FILE'

		if os.path.exists(XLSXfile) == True:
			XLSXfile_head, XLSXfile_tail = os.path.split(XLSXfile)
			print("Metrics file located.\nThe cleaned data will now be saved to the " + XLSXfile_tail + " file directly.")
			#Load the destination file into python
			mywb = load_workbook(XLSXfile)
			writer = pd.ExcelWriter(XLSXfile, engine='openpyxl') 
			writer.book = mywb
			writer.sheets = dict((ws.title, ws) for ws in mywb.worksheets)
			#Select the Data sheet and ignore the index column
			AllData.to_excel(writer, "Data", index = False)
			#write the data to the file
			writer.save()
		else:
			print("The " + XLSXfile_tail + " file could not be located in " + XLSXfile_head + ".")
	else:
		file_head, file_tail = os.path.split(file)
		print("The file " + file_tail + " could not be located at " + file_head + ".")


Metrics()

print ('This took '+ str(datetime.now()-startTime) + ' seconds')
time.sleep(7)

