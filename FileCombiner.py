# -*- coding: utf-8 -*-
"""
Created on Sat Apr  3 22:03:55 2021

@author: Ray Justin Huang
"""

# File Combiner
# A program to combine the data from multiple Excel files into one file

# Import necessary libraries
import pandas as pd
import glob2
import os

# Use two arguments from the command line: the input path and the output file
input_path = os.getcwd()
output_file = 'output_file/Consolidated_Data.xlsx'
sheet_name = 'consolidated_data'

# Collect filenames of all Excel sheets in the input path directory
all_workbooks = glob2.glob(os.path.join(input_path, '*.xls*'))

# Create subdirectory for output file
output_filedir = os.path.join(input_path,'output_file')
if not os.path.exists(output_filedir):
    os.mkdir(output_filedir)

# Collect data from all workbooks
data_list = []

for workbook in all_workbooks:
    all_worksheets = pd.read_excel(workbook, sheet_name=None, index_col=None)
    for worksheet_name, data in all_worksheets.items():
        data_list.append(data)

# Concatenate all data with pandas        
all_data_concatenated = pd.concat(data_list, axis=0, ignore_index=True)

# Write data to Excel file
with pd.ExcelWriter(output_file) as writer:
    all_data_concatenated.to_excel(writer, sheet_name=sheet_name, index=False)