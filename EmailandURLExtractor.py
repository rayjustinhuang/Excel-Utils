# -*- coding: utf-8 -*-
"""
Created on Sun Apr  4 21:02:33 2021

@author: Ray Justin Huang
"""

# Email and URL Extractor
# A program to extract emails and URLs from multiple Excel files

# Import necessary libraries
import pandas as pd
import glob2
import os
import itertools

# Use two arguments from the command line: the input path and the output file
input_path = os.getcwd()
output_file = 'output_file/All_Emails_and_URLs.xlsx'
sheet_name_emails = 'email_list'
sheet_name_urls = 'url_list'

# Collect filenames of all Excel sheets in the input path directory
all_workbooks = glob2.glob(os.path.join(input_path, '*.xls*'))

# Create subdirectory for output file
output_filedir = os.path.join(input_path,'output_file')
if not os.path.exists(output_filedir):
    os.mkdir(output_filedir)

# Collect emails and URLs from all workbooks
email_dict = {}
URL_dict = {}

for workbook in all_workbooks:
    all_worksheets = pd.read_excel(workbook, sheet_name=None, index_col=None)
    for worksheet_name, data in all_worksheets.items():
        for col in data.columns:
            workbook_name = workbook.split('\\')[-1]
            key = workbook_name + " | " + worksheet_name
            if data[col].dtypes == 'object':
                email_data = data[col].str.extractall(r'(.+@.+\.[a-zA-Z]{2,3})').values
                url_data = data[col].str.extractall(r'([Hh][Tt][Tt][Pp][Ss]?:\/\/[a-zA-Z0-9]+\.[a-zA-Z]{2,3}(\/[a-zA-Z0-9]+)*\.[a-zA-Z]{2,4})').values
                if len(email_data) > 0:
                    email_dict[key] = \
                        [str(i) for i in list(itertools.chain(*email_data))]
                if len(url_data) > 0:
                    URL_dict[key] = \
                        [str(i) for i in list(itertools.chain(*url_data))[::2]]

# Create dataframes for both emails and URLs
email_df = pd.DataFrame(email_dict)
URL_df = pd.DataFrame(URL_dict)

# Write data to Excel file
with pd.ExcelWriter(output_file) as writer:
    email_df.to_excel(writer, sheet_name=sheet_name_emails, index=False)
    URL_df.to_excel(writer, sheet_name=sheet_name_urls, index=False)

# Scratch work
# flat_email_list = [str(*i) for i in list(itertools.chain(*email_list))]
