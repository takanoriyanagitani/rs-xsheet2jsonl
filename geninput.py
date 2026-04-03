import xlsxwriter
import os
from datetime import datetime

ifile = os.getenv('ifile', './sample.d/input.xlsx')
os.makedirs(os.path.dirname(ifile), exist_ok=True)

workbook = xlsxwriter.Workbook(ifile)
worksheet = workbook.add_worksheet('Sheet1')

# Writing some sample data
data = [
    ['Header1', 'Header2', 'Header3'],
    [10, 20, 30],
    ['Row2', 12.34, True],
    ['Date', '2023-10-27', 'Wait...'],
]

for r, row in enumerate(data):
    for c, val in enumerate(row):
        worksheet.write(r, c, val)

# Specific date writing with format
date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
worksheet.write_datetime(3, 1, datetime(2023, 10, 27, 12, 0, 0), date_format)

workbook.close()
