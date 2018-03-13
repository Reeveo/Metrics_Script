import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import time


startTime = datetime.now()
#Get the raw data
raw_data = pd.read_csv('FILE PATH TO RAW DATA', header=0, encoding='ISO-8859-1', low_memory= False)


# This is the location of the destination file
file = ('FILE PATH OF THE XLSX FILE YOU WANT TO UPDATE')

#Load the destination file into python
mywb = load_workbook(file)
writer = pd.ExcelWriter(file, engine='openpyxl') 
writer.book = mywb
writer.sheets = dict((ws.title, ws) for ws in mywb.worksheets)

'''
Pandas
'''
#Remove any entries which have an error status
AllData = raw_data[~raw_data['Status'].str.contains('Error', na=False)==True]

#Select only the columns needed
AllData = pd.DataFrame(AllData, columns=[STATE THE COLUMNS YOU WANT TO KEEP])

#Select only cases which have specific string in the Assigned to Column
AllData = AllData[AllData['Assigned to'].str.contains('STATE THE STRING YOU WANT TO BE FOUND')==True]

#Format a date column to date time otherwise it will paste it as text and the formulas will not work
AllData[DATE COLUMN NAME] = pd.to_datetime(AllData['DATE COLUMN NAME'], dayfirst=True)

#Select the Data sheet and ignore the index column
AllData.to_excel(writer, "Data", index = False)

#write the data to the file
writer.save()
print ('This took '+str(datetime.now()-startTime) + ' seconds')
time.sleep(3)

