import xlsxwriter
from pathlib import Path

'''
This script's purpose is to convert a CSV file into an Microsfot Excell  spreadsheet. Takes in input the csv file's location; then
foreach row write the comma separated values in a column of the spreadsheet.

Uses XlsxWriter python module (see it's documentation).

This software is licenced as per Creative Commons BY-NC-SA 4.0, please refer to https://creativecommons.org/licenses/by-nc-sa/4.0/
for more informations.
'''

__author__ = "Sukhdev Mohan"
__version__ = "1.0.0"
__licence__ = "CC BY-NC-SA 4.0"

# Input csv_file
csv_file = Path(input("Please insert CSV file location: "))

# Insert output file
output_file = str(input("Please insert output file name without extension: ")) + '.xlsx'

with open(csv_file, 'r') as file:
    rows = [line.split(',') for line in file.read().splitlines()]

# Create column headers
# Static Column, please change the columns if the CSV file is changd
headers = ['BLOCK','TYPE','TIME','DATE','COUNT','CENTER X','CENTER Y','DIAMETER','LENGTH X','LENGTH Y','POSITION X','POSITION Y','POSITION Z','CORNER X','CORNER Y']
headers = [{'header':  header} for header in headers]

# Ddynamically define width of the table to write 
# Change destination letter ('o' in this case) if the columns change
table_def = 'A1:O' + str(len(rows))

# Create Excell object
workbook = xlsxwriter.Workbook(output_file)

# Add worksheet to the excell object
worksheet = workbook.add_worksheet()

# Write data
worksheet.add_table(table_def, {'data': rows, 'columns': headers})

# Save & close file
workbook.close()
