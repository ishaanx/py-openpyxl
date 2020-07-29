import pandas as pd
from openpyxl import *
from openpyxl.styles import *
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from time import sleep
import string
from alive_progress import alive_bar

with alive_bar(5, manual=True,title='Title here',theme='smooth',bar='blocks') as bar:   # default setting

	## Report 1 - 
	#convert csv to xlsx using pandas lib
	read_file = pd.read_csv (r'/Users/ishan/Documents/test.csv')
	bar (.10)
	read_file.to_excel (r'/Users/ishan/Documents/test.xlsx', index = None, header=True)
	bar(.20)
	#read xlsx
	#assign the excel file to wb() variable
	wb=load_workbook("/Users/ishan/Documents/test.xlsx")
	bar(.30)

	#assign the worksheet of the workbook to a ws() variable
	ws=wb.active
	bar(.40)
	mr = ws.max_row
	mc = ws.max_column
	for cell in ws[mr:mc]:
		cell.font = Font(size=11)
	bar(.60)
	# Set header row style
	for cell in ws["1:1"]:
		cell.font = Font(size=12)
		cell.style = 'Accent1'

	bar(.80)
	#set column width to 15 with loop
	for col in range(1, 54):
		ws.column_dimensions[(get_column_letter(col))].width = 15


	#Save the excel file
	bar(1)
	wb.save("/Users/ishan/Documents/test.xlsx")

	bar()
	#print ('Report Completed')




