import os
import xlsxwriter
import glob
import csv
import shutil
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from time import sleep
import string
from alive_progress import alive_bar
import sys
import subprocess
import time
from openpyxl import *
from openpyxl.styles import *
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook


v_chg_and_adj = input('Enter the file name for Charges and adjustments: ')
print(v_chg_and_adj)

def payments():

    print('Running one function')
    filehandler_path = open('Payments-Jun.csv', 'r')
    print(filehandler_path)
    output_path='./temp'

    def cleanup():
        print('Cleaning Up temp files')
        if os.path.exists(output_path):
            shutil.rmtree(output_path)


    def split(filehandler=filehandler_path, delimiter=',', row_limit=300000, 
        output_name_template='Payments_%s.csv',  keep_headers=True):
        """
        Splits a CSV file into multiple pieces.
        Arguments:
            `row_limit`: The number of rows you want in each output file. 10,000 by default.
            `output_name_template`: A %s-style template for the numbered output files.
            `output_path`: Where to stick the output files.
            `keep_headers`: Whether or not to print the headers in each output file.
        """
        print ('Creating Export paths')
        if not os.path.exists(output_path):
            os.makedirs(output_path)
        if not os.path.exists('Export'):
            os.makedirs('Export')

        reader = csv.reader(filehandler, delimiter=delimiter)
        current_piece = 1
        current_out_path = os.path.join(
             output_path,
             output_name_template  % current_piece
        )

        current_out_writer = csv.writer(open(current_out_path, 'w'), delimiter=delimiter)
        current_limit = row_limit
        if keep_headers:
            headers = next(reader)
            current_out_writer.writerow(headers)
        for i, row in enumerate(reader):
            if i + 1 > current_limit:
                current_piece += 1
                current_limit = row_limit * current_piece
                current_out_path = os.path.join(
                   output_path,
                   output_name_template  % current_piece
                )
                current_out_writer = csv.writer(open(current_out_path, 'w'), delimiter=delimiter)
                if keep_headers:
                    current_out_writer.writerow(headers)
            current_out_writer.writerow(row)
        return filehandler


    # join multiple csv to a single workbook with diff worksheets
    def split_join():
        print ('Joining files')
        # Returns the same day of last month if possible otherwise end of month
        # (eg: March 31st->29th Feb an July 31st->June 30th)
        last_month = datetime.now() - relativedelta(months=1)
        # Create string of month name and year...
        text = format(last_month, '%B %Y')
        prev_mon = '['+text+']'

        #print(prev_mon)
        fname = './Export/'+'Payments'+prev_mon+'.xlsx'
        writer = pd.ExcelWriter(fname, engine='xlsxwriter')
        folders = next(os.walk('.'))[1]
        for host in folders:
            Path = os.path.join(os.getcwd(), host)
            for f in glob.glob(os.path.join(Path, "Payments_*.csv")):
                #print(f)
                df = pd.read_csv(f,sep="\t",low_memory=False)
                df.to_excel(writer, index=False,sheet_name=os.path.basename(f)[:31])
            writer.save()
        return fname



    def styl():
        import openpyxl
        from openpyxl.reader.excel import load_workbook
        from openpyxl.styles import Font, Alignment 
        from openpyxl.utils import get_column_letter
        print('Applying')
        fname = split_join()
        print(fname)
        wb = load_workbook(fname)
        for ws in wb.worksheets:
            mr = ws.max_row
            mc = ws.max_column
            for cell in ws[mr:mc]:
                cell.font = Font(size=11)
            for cell in ws["1:1"]:
                cell.font = Font(size=12)
                cell.style = 'Accent1'
                cell.alignment = Alignment(wrapText='True',horizontal='center')
            for col in range(1, 54):
                ws.column_dimensions[(get_column_letter(col))].width = 15
                ws.freeze_panes = "A2"
        wb.save(fname)
    cleanup()
    split()
    split_join()
    styl()
    cleanup()



def chg_and_adj():
    print('Processing Charges and Adjustments Report')
    ## DECLARE VARIABLES
    #Source vars
    fs_path = os.getcwd()
    fs_name = v_chg_and_adj
    fs_ext = ".csv"


    ##Create working directory
    wd_name = "Export"
    wd = os.getcwd()
    #exp_dir=os.mkdir(wd_name)
    if not os.path.exists(wd_name):
        os.makedirs(wd_name)


    #Dest vars
    #fd_path = fs_path
    #fd_name = wd_name
    #fd_ext = ".xlsx"
    # Returns the same day of last month if possible otherwise end of month
    # (eg: March 31st->29th Feb an July 31st->June 30th)
    last_month = datetime.now() - relativedelta(months=1)
    # Create string of month name and year...
    text = format(last_month, '%B %Y')
    prev_mon = '['+text+']'

    #print(prev_mon)
    fd_name = './Export/'+'Charges and Adjustments'+prev_mon+'.xlsx'

    #main prog
    with alive_bar(total=100, manual='True',title='Report 1',theme='smooth',bar='blocks') as bar:   # default setting

        ## Report 1 - 
        #convert csv to xlsx using pandas lib
        read_file = pd.read_csv (''r''+fs_path+'/'+fs_name+fs_ext,sep="\t")
        bar(.10)
        read_file.to_excel (''r''+fd_name, index = None, header=True)
        bar(.20)
        #read xlsx
        #assign the excel file to wb() variable
        wb=load_workbook(fd_name)
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
            cell.alignment = Alignment(wrapText='True',horizontal='center')
        bar(.80)
        #set column width to 15 with loop
        for col in range(1, 54):
            ws.column_dimensions[(get_column_letter(col))].width = 15
            ws.freeze_panes = "A2"

        #Save the excel file
        bar(1)
        wb.save(fd_name)
        bar()
    print('Export Completed')


def cct():
    print('This is the 3 func')
def discp_rates():
    print('This is the 3 func')
def dnr1():
    print('This is the 3 func')
def dnr2():
    print('This is the 3 func')
def dnr3():
    print('This is the 3 func')
def grts_and_gst():
    print('This is the 3 func')
def gst_email():
    print('This is the 3 func')
def lldb():
    print('This is the 3 func')
def pay_and_ref():
    print('This is the 3 func')
def prop_over():
    print('This is the 3 func')
def room_moves():
    print('This is the 3 func')


def all():
    print('Processing all reports')
    payments()
    cct()
    chg_and_adj()
    discp_rates()
    dnr1()
    dnr2()
    dnr3()
    grts_and_gst()
    gst_email()
    lldb()
    pay_and_ref()
    prop_over()
    room_moves()
    

dispatcher = {
    '1': payments,
    '2': cct,
    '3': chg_and_adj,
    '4': discp_rates,
    '5': dnr1,
    '6': dnr2,
    '7': dnr3,
    '8': grts_and_gst,
    '9': gst_email,
    '10': lldb,
    '11': pay_and_ref,
    '12': prop_over,
    '13': room_moves,
    'all': all
}
print('Following choices are available: \n \
    1 - function one, \n\
    2 - function two , \n\
    3 - function three, \n\
    4 - run all four, \n\
    5 - function five, \n\
    6 - function six, \n\
    7 - function seven, \n\
    8 - function eigh, \n\
    9 - function nine, \n\
    10 - function ten, \n\
    11 - function eleven, \n\
    12 - function twelve, \n\
    13 - function thirteen, \n\
    ')

action = input('Option: - ')

dispatcher[action]()
