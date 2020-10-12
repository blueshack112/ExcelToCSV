#!/usr/bin/env python
# coding: utf-8

'''
    Transform Titan Excel Duplicates Feed file to a Flat CSV with 3 columns
    Output Field List
    - TireSize (Would be the value in Column 'A'
    - PartNumber (Would be the NON-Blank value for the current column that we are spinning through)
    - Tread (The Tread value would be the first part of the Column Name)
    
    Author: Patrick M Mahoney
    Date 2020 October 11
'''

# Need python version 3.4 or higher for pathlib
from pathlib import Path
import csv
import os
# import os.path
# from os import path
import pandas as pd

# Declare variables using Path from patlib
inputfile=Path.home() / "OneDrive - Mike's Auto Parts" / 'Share' / 'Titan Duplicate LT Passenger Crossover Sizes.xlsx'
inputsheet='Sheet1'
outputfile=Path.home() / 'Downloads' / 'Titan Fitment.tsv'

# Declare list of columns to specify formats from file to import
my_columns={
    'Light Tire Size':str,
    'Car-10xx':str,
    'Light Truck-10xx':str,
    'Car-15xx':str,
    'Light Truck-15xx':str,
    'Car-20xx':str,
    'Light Truck-20xx':str,
    'Car-23xx':str,
    'Light Truck-23xx':str,
    'Car-25xx':str,
    'Light Truck-25xx':str,
    'Car-30xx':str,
    'Light Truck-30xx':str,
    'Car-DC':str,
    'Light Truck-DC':str,
    'Car-Sock':str,
    'Light Truck-Sock':str
    }

# Validate file exists and available to open
# This doesn't work if the file is already opened by Excel
if os.path.exists(inputfile):
    try:
        os.rename(inputfile, inputfile)
        # print('File ' + str(inputfile) + ' Exists!')
        data = pd.read_excel(inputfile, sheet_name=inputsheet, converters=my_columns) # Read Excel file into python data frame
    except OSError as e:
        print('File Exists and Alredy OPEN by someone else')
else:
    print('False')

# Test for Null values
def isNaN(num):
          return num != num
        
# Returns the begining part for the column name before the hyphen
def name(word):
    new_word = word.split("-")
    return new_word[0]

# Spin through rows, then columns
row_list = []
for i in range(1,len(data.index)):
    for j in range(1, len(data.columns)-1):
        if not isNaN(data.iloc[i,j]):
            row_list.append({'TireSize': data.iloc[i,0], 'Part': data.iloc[i,j], 'Tread': name(data.columns[j])})
            
# Create data frame using row_list 
data_output = pd.DataFrame(row_list)

# Select list of fields to save to 
my_list = data_output.columns # or use my_list = ['Tire Size', 'Part', 'Tread']
                
# Write data frame by selected columns to csv file
# data.to_csv(outputfile, encoding='utf-8', escapechar='\\', float_format='%.2f', index=False, columns = my_list, quoting=csv.QUOTE_NONE, sep='\t') # Create csv file for SQL Server to import
data_output.to_csv(outputfile, encoding='utf-8', escapechar='\\', index=False, quoting=csv.QUOTE_NONE, sep='\t' , columns = my_list)
'''
    columns = my_list      - Only save selected columns from my_list
    encoding='utf-8'       - Use utf encoding
    float_format='%.2f'    - Set to 2 decimal places
    index=False            - Turn off row number
    quoting=csv.QUOTE_NONE - Don't surround text columns with double quotes
    sep=','                - Use comma as column delimiter
'''
