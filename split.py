import os
import xlsxwriter
import glob
import csv
import shutil
import pandas as pd
import xlsxwriter
from datetime import datetime
from dateutil.relativedelta import relativedelta




output_path='./temp'

def cleanup():
    if os.path.exists(output_path):
        shutil.rmtree(output_path)

def split(filehandler, delimiter=',', row_limit=300000, 
    output_name_template='Payments_%s.csv',  keep_headers=True):
    """
    Splits a CSV file into multiple pieces.
    
    A quick bastardization of the Python CSV library.
    Arguments:
        `row_limit`: The number of rows you want in each output file. 10,000 by default.
        `output_name_template`: A %s-style template for the numbered output files.
        `output_path`: Where to stick the output files.
        `keep_headers`: Whether or not to print the headers in each output file.
    Example usage:
    
        >> from toolbox import csv_splitter;
        >> csv_splitter.split(open('/home/ben/input.csv', 'r'));
    
    """
    if not os.path.exists(output_path):
        os.makedirs(output_path)

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



# join multiple csv to a single workbook with diff worksheets
def split_join():
        # Returns the same day of last month if possible otherwise end of month
    # (eg: March 31st->29th Feb an July 31st->June 30th)
    last_month = datetime.now() - relativedelta(months=1)

    # Create string of month name and year...
    text = format(last_month, '%B %Y')
    prev_mon = '['+text+']'
    print(prev_mon)
    fname = 'Payments'+prev_mon+'.xlsx'
    writer = pd.ExcelWriter(fname, engine='xlsxwriter')
    folders = next(os.walk('.'))[1]
    for host in folders:
        Path = os.path.join(os.getcwd(), host)
        for f in glob.glob(os.path.join(Path, "Payments_*.csv")):
            print(f)
            df = pd.read_csv(f,sep="\t",low_memory=False)
            df.to_excel(writer, index=False,sheet_name=os.path.basename(f)[:31])

    writer.save()
    

cleanup()
split(open('Payments-Jun.csv', 'r'))
split_join()
cleanup()




