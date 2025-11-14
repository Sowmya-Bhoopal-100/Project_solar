# This script performs data wrangling and creates reports from an iAuditor 'sqlite.db' file.
# Antonio Mantilla 2025

'''
iAuditor is part of Schneider Solar Services Safety Culture initiative https://app.safetyculture.com/home

Learn more about the template here: https://schneiderelectric.sharepoint.com/sites/SolarServicesRDSustainingGroup/SitePages/Data-Analytics.aspx

Link to "SEW-721  iAuditor template SE-Global Customer Activity Report - COMMENTS.xlsx" file: 
https://schneiderelectric.sharepoint.com/:x:/r/sites/SolarServicesRDSustainingGroup/Shared%20Documents/SEW%20Projects%20Folders/SEW-719%20iAuditor%20data%20analysis/SEW-721%20%20iAuditor%20template%20SE-Global%20Customer%20Activity%20Report%20-%20COMMENTS.xlsx?d=w83327c1da9f9494cac8df509f60dfdb1&csf=1&web=1&e=pHQErR

This program takes the csv file exported from iAuditor and implements the following:
1. Add new column that combines the question category with the question title.
2. Pivot data frame so that each question becomes a column and the answers becomes the values in the column. 
The resulting data frame has one row per report.
3. Create new directory and export the resulting data frame into a csv file.

It can also open a db file exported from iAuditor. 

Note: each iAuditor forms has a 'template_id'. For example, form "SE-Global Customer Activity Report" has template_id = template_b99f63ecf11b4de0909cfb362952fb9a. 
The form can be open at https://app.safetyculture.com/template-editor/template_b99f63ecf11b4de0909cfb362952fb9a

For info about how to extract data from iAuditor, go to https://jira.se.com/browse/SEW-719

GitHub repository: https://github.schneider-electric.com/solar-services-data-analytics/customer_activity_report_data_wrangling/tree/main

'''

# To create executable file, run pyinstaller iAuditor_report_generator.spec or pyinstaller iAuditor_report_generator.py --onefile

# For help with tkinter GUI design: https://realpython.com/python-gui-tkinter/


import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog
from tkinter import StringVar
from tkinter import OptionMenu
from tkinter import Listbox
from tkinter import ttk
#import pdfrw

# Load the Pandas libraries with alias 'pd'
import pandas as pd 
import sys
import numpy as np
import os
import glob
import linecache
import pathlib
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import mplcursors
import seaborn as sns

#from fpdf import FPDF

from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, PageTemplate, Frame
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Image as ReportLabImage
from reportlab.lib.units import inch

from PIL import Image as PILImage

import time
from datetime import datetime
from datetime import date
import dateparser
import threading

import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.formatting.rule import CellIsRule,FormulaRule
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

import PyPDF2
import sqlite3
import re

# DEFINE CONSTANTS ************************************************************************************

VERSION = ' Version 1.9 - EXPERIMENTAL '
PROGRAM_TITLE = "iAuditor Export Report Generator"
INSTRUCTIONS = ("If file 'sqlit_inspection_items_dataframe.csv' has not been created yet, " 
"use the button 'Select db file...' to extract the csv file from the 'sqlite.db' file extracted from iAuditor. "
"\nOtherwise, load the 'sqlit_inspection_items_dataframe.csv' file to be analyzed using button 'Select file to analyze...'.\n")


# DEFINE FILE COLUMNS
AUDIT_ID_COLUMN = 'audit_id'  # Field report identification number
ITEM_INDEX_COLUMN ='item_index'
QUESTION_COLUMN = 'label'
ANSWER_COLUMN = 'response'
QUESTION_TYPE_COLUMN = 'type'
QUESTION_CATEGORY_COLUMN = 'category'
QUESTION_COMBINED_LABEL_COLUMN = 'Question combined label'
INVERTER_SN_COLUMN = 'General Information - Inverter Serial Number'
INVERTER_MODEL_COLUMN = 'General Information - Model'
CASE_TYPE_COLUMN = 'General Information - Type of Service'
TECH_NAME_COLUMN = 'General Information - Technician Name*'
SITE_NAME_COLUMN = 'Site Information - Site Name*'
SERVICE_DATE_COLUMN = 'General Information - Service Date (YYYY-MM-DD)*'
SERVICE_DATE_FORMATTED_COLUMN = 'Service date formatted Y-m-d'
INVERTER_TECHNOLOGY_COLUMN = 'Inverter Preventive Actions - Checklist - Select Inverter technology'
PARENT_IDS_COLUMN = 'parent_ids'
ITEM_ID_COLUMN = 'item_id'
TYPE_COLUMN = 'type'



# today_date = datetime.today()
Total_number_pages = '?'
# # Get today's date
# today = datetime.today()
# # Format the date
# formatted_date = today.strftime('%d/%b/%Y')




def PrintException():
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    print('EXCEPTION IN ({}, LINE {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj))
    exception_Info = str('OPERATION INTERRUPTED. EXCEPTION IN ({}, LINE {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj))
    printToScreen(exception_Info)
    updateStatusBar("Error", True)


def printToScreen(the_text):
    txt_edit.insert(tk.END, str(the_text) + "\n" )
    txt_edit.yview(tk.END)
    txt_edit.update_idletasks()    

def printToScreen_with_timestamp(the_text):
    txt_edit.insert(tk.END, str(str(datetime.now()) + ' -> ' + the_text) + "\n" )
    txt_edit.yview(tk.END)
    txt_edit.update_idletasks()  


def updateStatusBar(message,warning):
    if warning == True:
        status_bar.config(bg= '#ded7a4', fg= 'black')
    else:
        status_bar.config(bg= default_bg, fg= default_fg)

    status_bar.config(text=f'Status: {message}')
    status_bar.update_idletasks()    

def select_input_file():
    filepath = askopenfilename(initialdir="", title="Select file ",
        defaultextension="csv",
        filetypes=[("Comma Delimited Files", ".csv")],
    )
    
    return filepath

def select_output_file(default_file_name, default_output_folder):

    # printToScreen("Please select output file.")    
    filepath = asksaveasfilename(initialdir=default_output_folder, title="Select output file",initialfile = default_file_name,
        defaultextension="csv",
        filetypes=[("Comma Delimited Files", "*.csv")],
    )

    return filepath


def get_device_data(the_data_frame):
    """ 
    The iAuditor form collects information about devices used during the intervention. 
    There are 3 fields devices (listed below where [nn] is the index of the part. The
    first device is 1, the second device is 2, and so on). 
        Device [nn] - Indicate type:
        Device [nn] - Serial Number
        Device [nn] - Type of tool

    This function extract the information about devices by AUDIT_ID_COLUMN.

    """ 


    try:

        # create data frame that only contains AUDIT_ID_COLUMN and any column that starts with 'Device'
        filtered_columns = [AUDIT_ID_COLUMN] + [col for col in the_data_frame.columns if col.startswith('Device')]
        filtered_df = the_data_frame[filtered_columns]

        # Sort DataFrame by column names
        filtered_df = filtered_df.sort_index(axis=1)
        print(filtered_df.shape)
        # Eliminate rows where the only column with values is AUDIT_ID_COLUMN
        filtered_df = filtered_df.dropna(how='all', subset=filtered_df.columns.difference([AUDIT_ID_COLUMN]))
        print(filtered_df.shape)

        # Eliminate columns that have no values
        filtered_df = filtered_df.dropna(axis=1, how='all')
        print(filtered_df.shape)

        # get the highest number that appears in the columns names
        numeric_values = []
        for name in filtered_columns:
            match = re.search(r'\d+', name)
            if match:
                numeric_values.append(int(match.group()))

        # Identify the highest number
        highest_number = max(numeric_values) if numeric_values else None
        print(highest_number)
        # Columns to be created
        new_columns = [
            AUDIT_ID_COLUMN,
            "Device - Indicate type:",
            "Device - Serial Number",
            "Device - Type of tool"
        ]

        # Initialize an empty list to store the new rows
        new_rows = []


            
        # Iterate over each part data column set
        for index, row in filtered_df.iterrows():
            audit_id = row[AUDIT_ID_COLUMN]
            
            # Iterate over each device column set
            for i in range(1, highest_number):  
                new_row = {
                    AUDIT_ID_COLUMN: audit_id,
                    "Device - Indicate type:": row.get(f"Device {i} - Indicate type:"),
                    "Device - Serial Number": row.get(f"Device {i} - Serial Number"),
                    "Device - Type of tool": row.get(f"Device {i} - Type of tool")
                }
                new_rows.append(new_row)

        # Create a new DataFrame from the list of new rows
        new_df = pd.DataFrame(new_rows, columns=new_columns)
        new_df = new_df.set_index(AUDIT_ID_COLUMN)
        print(new_df.shape)

        # Count rows with all empty values
        empty_rows_count = new_df.isna().all(axis=1).sum()
        print(f"Number of rows with all empty values: {empty_rows_count}")
        new_df = new_df.dropna(how='all')
        # Reset the index to make AUDIT_ID_COLUMN a regular column again
        new_df = new_df.reset_index()
        print(new_df.shape)
        return new_df

    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        PrintException() 

def get_part_replace_data(the_data_frame):
    """ 
    The iAuditor form collects information about parts replaced during the intervention. 
    There are six fields per part (listed below where [nn] is the index of the part. The
    first part is 1, the second part is 2, and so on). 
        Part Data [nn] - Part Designator
        Part Data [nn] - Part Number
    	Part Data [nn] - Part Reference Designator (ex. PP601)
        Part Data [nn] - Quantity	
        Part Data [nn] - Serial number - NEW part	
        Part Data [nn] - Serial number - REPLACED part

    This function extract the information about parts replaced by AUDIT_ID_COLUMN.

    """ 


    try:

        # create data frame that only contains AUDIT_ID_COLUMN and any column that starts with 'Part Data'
        filtered_columns = [AUDIT_ID_COLUMN] + [col for col in the_data_frame.columns if col.startswith('Part Data')] + [SERVICE_DATE_COLUMN]
        filtered_df = the_data_frame[filtered_columns]

        # Sort DataFrame by column names
        filtered_df = filtered_df.sort_index(axis=1)
        print(filtered_df.shape)
        # Eliminate rows where the only column with values is AUDIT_ID_COLUMN
        filtered_df = filtered_df.dropna(how='all', subset=filtered_df.columns.difference([AUDIT_ID_COLUMN]))
        print(filtered_df.shape)

        # Eliminate columns that have no values
        filtered_df = filtered_df.dropna(axis=1, how='all')
        print(filtered_df.shape)

        # get the highest number that appears in the columns names
        numeric_values = []
        for name in filtered_columns:
            match = re.search(r'\d+', name)
            if match:
                numeric_values.append(int(match.group()))

        # Identify the highest number
        highest_number = max(numeric_values) if numeric_values else None
        print(highest_number)
        # Columns to be created
        new_columns = [
            AUDIT_ID_COLUMN,
            "Part Data - Part Designator",
            "Part Data - Part Number",
            "Part Data - Part Reference Designator (ex. PP601)",
            "Part Data - Quantity",
            "Part Data - Serial number - NEW part",
            "Part Data - Serial number - REPLACED part",
            SERVICE_DATE_COLUMN
        ]

        # Initialize an empty list to store the new rows
        new_rows = []

        printToScreen_with_timestamp("Starting part replacement checks...")
            
        # Iterate over each part data column set
        for index, row in filtered_df.iterrows():
            audit_id = row[AUDIT_ID_COLUMN]
            service_date = row[SERVICE_DATE_COLUMN]
            # Iterate over each part data column set
            for i in range(1, highest_number):  
            #    if pd.notna(row.get(f"Part Data {i} - Part Number")) and pd.notna(row.get(f"Part Data {i} - Part Quantity")):
                new_row = {
                    AUDIT_ID_COLUMN: audit_id,
                    "Part Data - Part Designator": row.get(f"Part Data {i} - Part Designator"),
                    "Part Data - Part Number": row.get(f"Part Data {i} - Part Number"),
                    "Part Data - Part Reference Designator (ex. PP601)": row.get(f"Part Data {i} - Part Reference Designator (ex. PP601)"),
                    "Part Data - Part Quantity": row.get(f"Part Data {i} - Part Quantity"),                    
                    "Part Data - Serial number - NEW part": row.get(f"Part Data {i} - Serial number - NEW part"),
                    "Part Data - Serial number - REPLACED part": row.get(f"Part Data {i} - Serial number - REPLACED part"),
                    SERVICE_DATE_COLUMN: service_date
                }
                if pd.isna(row.get(f"Part Data {i} - Part Number")) and pd.isna(row.get(f"Part Data {i} - Part Quantity")):
                    pass
                else:    
                    new_rows.append(new_row)                     


        printToScreen_with_timestamp("Completed part replacement checks!")


        # Create a new DataFrame from the list of new rows
        new_df = pd.DataFrame(new_rows, columns=new_columns)
        new_df = new_df.set_index(AUDIT_ID_COLUMN)
        print(new_df.shape)

        # Count rows with all empty values
        empty_rows_count = new_df.isna().all(axis=1).sum()
        print(f"Number of rows with all empty values: {empty_rows_count}")
        new_df = new_df.dropna(how='all')
        # Reset the index to make AUDIT_ID_COLUMN a regular column again
        new_df = new_df.reset_index()
        print(new_df.shape)
        return new_df

    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        PrintException() 

def open_db_file():

    try:
        filepath = askopenfilename(initialdir="", title="Select db file ",
            defaultextension="db",
            filetypes=[("Comma Delimited Files", ".db")],
        )
        if not filepath:
            print("Input file was not selected")
            return
        printToScreen("File selected: " + filepath)


        # Get file information
        # Get the file's stat information
        file_path_2 = Path(filepath)
        stat_info = file_path_2.stat()

        try:
            file_creation_time = stat_info.st_birthtime 
        except:
            file_creation_time = stat_info.st_mtime # In case operating system did not like st_birthtime

        # Convert to megabytes
        size_in_mb = stat_info.st_size / (1024 * 1024)
 
        # Format using time.strftime() for more control
        file_modified_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(stat_info.st_mtime))
        file_created_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(file_creation_time))

        printToScreen(f'File size: {size_in_mb:.2f} MB')
        printToScreen(f'Last modified: {file_modified_time}')
        printToScreen(f'Created: {file_created_time}') 


        # Connect to the SQLite database
        conn = sqlite3.connect(filepath)


        # Query to get the list of tables
        tables_query = "SELECT name FROM sqlite_master WHERE type='table';"
        tables_df = pd.read_sql_query(tables_query, conn)
        # Print the names of the tables
        print(tables_df)

        your_table_name = 'inspection_items'
        # Load data into a DataFrame
        query = f"SELECT * FROM {your_table_name}"  # Replace with your table name
        iAuditor_df = pd.read_sql_query(query, conn)

        # Close the connection
        conn.close()

        output_file_name = filepath[:-4] + '_' + your_table_name + "_dataframe.csv"
        iAuditor_df.to_csv(output_file_name)
        printToScreen(f"\nTable {your_table_name} from the db file has been converted into file: {output_file_name}.")
        print(iAuditor_df.shape)    

    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        PrintException() 


def Select_file_and_analysis():

    try:
        today_date = datetime.today()
        input_file = select_input_file()

        if not input_file:
            print("Input file was not selected")
            return
        printToScreen("File selected: " + input_file)
        
        # Get file information
        # Get the file's stat information
        file_path = Path(input_file)
        stat_info = file_path.stat()


        try:
            file_creation_time = stat_info.st_birthtime 
        except:
            file_creation_time = stat_info.st_mtime  # In case operating system did not like st_birthtime



        # Access specific attributes

        # Convert to megabytes
        size_in_mb = stat_info.st_size / (1024 * 1024)
 
        # Format using time.strftime() for more control
        file_modified_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(stat_info.st_mtime))
        file_created_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(file_creation_time))

        printToScreen(f'File size: {size_in_mb:.2f} MB')
        printToScreen(f'Last modified: {file_modified_time}')
        printToScreen(f'Created: {file_created_time}') 
        printToScreen("This file was created on: " + file_created_time)


        data_raw = pd.read_csv(input_file) 

        print(data_raw.shape)
        print(list(data_raw.columns))

        # Count the number of unique values in the column 'AUDIT_ID_COLUMN'
        reports_in_file_count = data_raw[AUDIT_ID_COLUMN].nunique()

        printToScreen(f"\nThere are {reports_in_file_count} inspection reports in the file.")
        printToScreen(f"Analyzing {data_raw.shape[0]*data_raw.shape[1]:,} data points.")        
        printToScreen_with_timestamp("\nCreating inspections database... This will take a few minutes...")

        start_time = datetime.now()
        updateStatusBar("Creating inspections database...",False)
       
        HEADER_2 = 'iAuditor Report: '

        # Format the date
        formatted_today_date= today_date.strftime('%d_%b_%Y_%H_%M')
        directory_name = formatted_today_date

        # Create a directory to save files and images if it doesn't exist
        cl_output_dir = file_path.parent / directory_name
        FILE_NAME_TEXT = ' Customer_Activity_Report'
        os.makedirs(cl_output_dir, exist_ok=True)   
        output_file_selected = os.path.join(cl_output_dir, 'iAuditor_' + FILE_NAME_TEXT + ".csv")



        create_iAuditor_report(data_raw, output_file_selected, cl_output_dir,file_created_time, HEADER_2)

        end_time = datetime.now()   
        execution_time = end_time - start_time
        # Convert execution time to total seconds
        total_seconds = execution_time.total_seconds()

        # Calculate hours, minutes, and seconds
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)

        # Print the formatted execution time based on the values of hours and minutes
        if hours > 0:
            printToScreen(f'Execution time: {int(hours)} hours, {int(minutes)} minutes, {seconds:.2f} seconds')
        elif minutes > 0:
            printToScreen(f'Execution time: {int(minutes)} minutes, {seconds:.2f} seconds')
        else:
            printToScreen(f'Execution time: {seconds:.2f} seconds')

        printToScreen_with_timestamp("\n\nANALYSIS COMPLETED!")
        updateStatusBar("ANALYSIS COMPLETED.",False)

    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        PrintException()        


# Define a function to create QUESTION_COMBINED_LABEL_COLUMN
def create_combined_label(row, this_data):
    try:

        if "Anomaly?" in str(row[QUESTION_COLUMN]): # second parent ID tells the question which is the parent
            # Split PARENT_IDS_COLUMN and get the second item (index 1)
            parent_ids = row[PARENT_IDS_COLUMN].split(',')
            if len(parent_ids) > 1:  # Check if there is a second item
                second_parent_id = parent_ids[1]
                # Find the QUESTION_COLUMN for the matching ITEM_ID_COLUMN
                parent_row = this_data[this_data[ITEM_ID_COLUMN] == second_parent_id]
                if not parent_row.empty:  # Check if the parent row exists
                    parent_question = parent_row[QUESTION_COLUMN].values[0]  # Get the QUESTION_COLUMN value
                    return f"{parent_question} - {row[QUESTION_COLUMN]}"
                else:       
                    return f"{second_parent_id} - {row[QUESTION_COLUMN]}"
        elif "if response is" in str(row[QUESTION_COLUMN]): # first parent ID tells the question which is the parent
            # Split PARENT_IDS_COLUMN and get the second item (index 1)
            parent_ids = row[PARENT_IDS_COLUMN].split(',')
            if len(parent_ids) > 1:  # Check if there is a second item
                first_parent_id = parent_ids[0]
                # Find the QUESTION_COLUMN for the matching ITEM_ID_COLUMN
                parent_row = this_data[this_data[ITEM_ID_COLUMN] == first_parent_id]
                if not parent_row.empty:  # Check if the parent row exists
                    parent_question = parent_row[QUESTION_COLUMN].values[0]  # Get the QUESTION_COLUMN value
                    return f"{parent_question} - {row[QUESTION_COLUMN]}"
                else:       
                    return f"{first_parent_id} - {row[QUESTION_COLUMN]}"
        else:
            return f"{row[QUESTION_CATEGORY_COLUMN]} - {row[QUESTION_COLUMN]}"

    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        PrintException()

# Define a function to create QUESTION_COMBINED_LABEL_COLUMN by PARENT_IDS_COLUMN
def create_combined_label_by_parentIDs(row, this_data):
    try:

        if "Anomaly?" in str(row[QUESTION_COLUMN]): # second parent ID tells the question which is the parent
            # Split PARENT_IDS_COLUMN and get the second item (index 1)
            parent_ids = row[PARENT_IDS_COLUMN].split(',')
            if len(parent_ids) > 1:  # Check if there is a second item
                second_parent_id = parent_ids[1]
                # Find the QUESTION_COLUMN for the matching ITEM_ID_COLUMN
                parent_row = this_data[this_data[ITEM_ID_COLUMN] == second_parent_id]
                if not parent_row.empty:  # Check if the parent row exists
                    parent_question = parent_row[QUESTION_COLUMN].values[0]  # Get the QUESTION_COLUMN value
                    return f"{parent_question} - {row[QUESTION_COLUMN]} - {row[ITEM_ID_COLUMN]}"
                else:       
                    return f"{second_parent_id} - {row[QUESTION_COLUMN]} - {row[ITEM_ID_COLUMN]}"
        elif "if response is" in str(row[QUESTION_COLUMN]): # first parent ID tells the question which is the parent
            # Split PARENT_IDS_COLUMN and get the second item (index 1)
            parent_ids = row[PARENT_IDS_COLUMN].split(',')
            if len(parent_ids) > 1:  # Check if there is a second item
                first_parent_id = parent_ids[0]
                # Find the QUESTION_COLUMN for the matching ITEM_ID_COLUMN
                parent_row = this_data[this_data[ITEM_ID_COLUMN] == first_parent_id]
                if not parent_row.empty:  # Check if the parent row exists
                    parent_question = parent_row[QUESTION_COLUMN].values[0]  # Get the QUESTION_COLUMN value
                    return f"{parent_question} - {row[QUESTION_COLUMN]} - {row[ITEM_ID_COLUMN]}"
                else:       
                    return f"{first_parent_id} - {row[QUESTION_COLUMN]} - {row[ITEM_ID_COLUMN]}"
        else:
            return f"{row[QUESTION_CATEGORY_COLUMN]} - {row[QUESTION_COLUMN]} - {row[ITEM_ID_COLUMN]}"

    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        PrintException()

def determine_auditID_by_year(dataframe, year):

    SERVICE_DATE_LABEL = 'Service Date (YYYY-MM-DD)*'
    
    
    try:
        
        data_in_year = dataframe[dataframe[QUESTION_COLUMN] == SERVICE_DATE_LABEL] 
        print(data_in_year.shape)
        # data_in_year[ANSWER_COLUMN] = pd.to_datetime(data_in_year[ANSWER_COLUMN], errors='coerce')
        data_in_year[ANSWER_COLUMN] = data_in_year[ANSWER_COLUMN].apply(parse_datetime)
        
       # data_in_year.to_csv('datetime_transform.csv') # for debugging

        # Count the number of NaT values in ANSWER_COLUMN
        num_invalid_dates = data_in_year[ANSWER_COLUMN].isna().sum()
        print(f'num_invalid_dates: {num_invalid_dates}')
        print(data_in_year.shape)
        # Drop rows with NaT values in ANSWER_COLUMN
        data_in_year = data_in_year.dropna(subset=[ANSWER_COLUMN])
        print(data_in_year.shape)
        data_in_year = data_in_year[data_in_year[ANSWER_COLUMN].dt.year == year]
        print(data_in_year.shape)
        unique_audit_ids = data_in_year[AUDIT_ID_COLUMN].unique()
        
        return unique_audit_ids
        
    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        PrintException()

def parse_datetime(date_str):
    # required because The answer for "Service Date (YYYY-MM-DD)*" have different formats. Sometimes is 2024-11-12T14:01:35Z and others 2024-10-11T05:58:03.949Z
    try:
        return pd.to_datetime(date_str, format='%Y-%m-%dT%H:%M:%SZ', errors='raise')
    except ValueError:
        return pd.to_datetime(date_str, format='%Y-%m-%dT%H:%M:%S.%fZ', errors='coerce')


def create_iAuditor_report(this_data, output_file, output_dir, this_file_created_time, header_2):
    try:
        data = this_data
        output_file_selected = output_file
        cl_output_dir = output_dir
        file_created_time = this_file_created_time
        HEADER_2 = header_2

 
        number_of_records = data.shape[0]
        # remove records of type = 'information' as they don't have data. This is to reduce number of unnecessary columns when creating QUESTION_COMBINED_LABEL_COLUMN
        data = data[data[TYPE_COLUMN] != 'information']
        print(data.shape)
        printToScreen("Records of type 'information' have been removed.")
        data = data.dropna(subset=[QUESTION_COLUMN]) # if the question is blank, there is no valid answer. 
        printToScreen("Records without data in column 'label' have been removed.")
        print(data.shape)
        data = data[data[TYPE_COLUMN] != 'section']
        print(data.shape)
        printToScreen("Records of type 'section' have been removed.")  
        data = data[data[TYPE_COLUMN] != 'signature']
        print(data.shape)
        printToScreen("Records of type 'signature' have been removed.")  

        # Count duplicates
        duplicate_count = data.duplicated().sum()
        # Remove duplicates
        data = data.drop_duplicates()
        print(data.shape)
        printToScreen(f"{duplicate_count} Duplicate records have been removed.")  

        number_of_records_removed = number_of_records - data.shape[0]   
        printToScreen(f'Total number of records removed: {number_of_records_removed:,}') 

        printToScreen_with_timestamp("\nData wrangling in process...it will take a few minutes...")  
        updateStatusBar("Data wrangling in process...",False)

        # Add a new column that combines the values of QUESTION_CATEGORY_COLUMN and QUESTION_COLUMN joined by " - "
        # data[QUESTION_COMBINED_LABEL_COLUMN] = data[QUESTION_CATEGORY_COLUMN] + " - " + data[QUESTION_COLUMN]
        # Apply the function to create the new column

        data[QUESTION_COMBINED_LABEL_COLUMN] = data.apply(create_combined_label, axis=1, args=(data,))           


        updateStatusBar("Building output files...",False)  
        first_output_file = output_file_selected[:-4] +  "_combined_label.csv"
        data = data.sort_values(by=[AUDIT_ID_COLUMN, ITEM_INDEX_COLUMN])
        data.to_csv(first_output_file, index=False)
        printToScreen_with_timestamp(f"\nRaw data with added column {QUESTION_COMBINED_LABEL_COLUMN} has been exported to file: " + first_output_file + "\n")
        printToScreen(f'This file is sorted by {AUDIT_ID_COLUMN} and {ITEM_INDEX_COLUMN}')


        # Pivot the dataframe to transform QUESTION_COMBINED_LABEL_COLUMN values into columns while keeping AUDIT_ID_COLUMN
        pivoted_df = data.pivot_table(index=AUDIT_ID_COLUMN, columns=QUESTION_COMBINED_LABEL_COLUMN, values=ANSWER_COLUMN, aggfunc='first')

        # Reset the index to make it a proper dataframe
        pivoted_df.reset_index(inplace=True)

        # Sort the resulting dataframe by AUDIT_ID_COLUMN
        sorted_df = pivoted_df.sort_values(by=AUDIT_ID_COLUMN)

        # extract the information about parts replaced
        printToScreen_with_timestamp("\nExtracting parts replaced data...")
        updateStatusBar("Extracting parts replaced data...",False)        
        parts_replaced_df = get_part_replace_data(sorted_df)
        parts_replaced_file_name = output_file_selected[:-4] + "_PartsReplaced.csv"
        parts_replaced_df.to_csv(parts_replaced_file_name, index=False)

        printToScreen("Parts replaced data have been extracted into file: " + parts_replaced_file_name)
        printToScreen_with_timestamp("\nParts replace extraction completed!")
        printToScreen("Removing columns that start with 'Part Data'...")
        # remove columns that start with 'Part Data'
        sorted_df = sorted_df.loc[:, ~sorted_df.columns.str.startswith('Part Data')]


        # extract the information about devices
        printToScreen_with_timestamp("\nExtracting devices data...")
        updateStatusBar("Extracting devices data...",False)        
        devices_df = get_device_data(sorted_df)
        devices_file_name = output_file_selected[:-4] + "_devices.csv"
        devices_df.to_csv(devices_file_name, index=False)

        printToScreen("Devices data have been extracted into file: " + devices_file_name)
        printToScreen_with_timestamp("\nDevices extraction completed!")
        printToScreen("Removing columns that start with 'Device'...")
        # remove columns that start with 'Device'
        sorted_df = sorted_df.loc[:, ~sorted_df.columns.str.startswith('Device')]


        # Add date column
        sorted_df[SERVICE_DATE_COLUMN] = pd.to_datetime(sorted_df[SERVICE_DATE_COLUMN], errors='coerce')
        # Add a new column with the date in "%Y-%m-%d" format
        sorted_df[SERVICE_DATE_FORMATTED_COLUMN] = sorted_df[SERVICE_DATE_COLUMN].dt.strftime('%Y-%m-%d')
        printToScreen(f"\n Column '{SERVICE_DATE_FORMATTED_COLUMN}' added.")



        sorted_df.to_csv(output_file_selected, index=False)

        printToScreen("\n File with one row per inspection and inspection questions as columns has been created: " + output_file_selected + "\n")

        printToScreen(f"\nNumber of records: {sorted_df.shape[0]}.")

        
       
        updateStatusBar("Creating sqlite database.",False)
        # *********** CREATE SQLITE DATABASE **************************************
        printToScreen('\nCreating sqlite database...')
        database_output_file = output_file_selected[:-4] +  "_database.db"    
        # Connect to SQLite database (or create it if it doesn't exist)
        conn = sqlite3.connect(database_output_file)



        # # Check for duplicate columm names that can cause an exception when converting to_sql
        # for column in sorted_df.columns:
        #     printToScreen(column)
        # duplicate_columns = sorted_df.columns[sorted_df.columns.duplicated()].tolist()
        # printToScreen(f"Duplicate columns: {duplicate_columns}")

        # Convert DataFrames to SQL table
        try:
            sorted_df.to_sql('main_table', conn, if_exists='replace', index=False)
        except Exception as e:
            print("Oops!", e.__class__, "occurred.")
            PrintException()

        try:
            parts_replaced_df.to_sql('replaced_parts', conn, if_exists='replace', index=False)
        except Exception as e:
            print("Oops!", e.__class__, "occurred.")
            PrintException()       
        
        try:
            devices_df.to_sql('devices', conn, if_exists='replace', index=False)
        except Exception as e:
            print("Oops!", e.__class__, "occurred.")
            PrintException()

        # Close the connection
        conn.close()
        printToScreen("\nSQL database file is: " + database_output_file + "\n")

        # ANALIZE THE DATA IN THE FILE
        printToScreen('\n************ SOME DATA ANALYSIS ******************')
        do_column_overview(INVERTER_SN_COLUMN,sorted_df)     
        do_column_overview(INVERTER_TECHNOLOGY_COLUMN,sorted_df)             
        do_column_overview(INVERTER_MODEL_COLUMN,sorted_df)
        do_column_overview(CASE_TYPE_COLUMN,sorted_df)
        do_column_overview(TECH_NAME_COLUMN,sorted_df)        


        do_column_overview(SITE_NAME_COLUMN,sorted_df)
       # non_numeric_count = sorted_df.dropna(subset=[SITE_NAME_COLUMN])[~sorted_df[SITE_NAME_COLUMN].str.split().str[0].str.isnumeric()].shape[0]
        non_numeric_count = sorted_df.dropna(subset=[SITE_NAME_COLUMN])
        non_numeric_count = non_numeric_count[non_numeric_count[SITE_NAME_COLUMN].str.split().str[0].str.isnumeric() == False].shape[0]
        printToScreen(f"\nNumber of records without a site number: {non_numeric_count}")

       # ANALIZE SERVICE DATE 
     #   sorted_df[SERVICE_DATE_COLUMN] = pd.to_datetime(sorted_df[SERVICE_DATE_COLUMN], errors='coerce')
        do_column_overview(SERVICE_DATE_COLUMN,sorted_df) 
        # Count NaT instances in the SERVICE_DATE_COLUMN
        nat_count = sorted_df[SERVICE_DATE_COLUMN].isna().sum()
        printToScreen(f"Number of NaT instances: {nat_count}")

        sorted_df['YearMonth'] = sorted_df[SERVICE_DATE_COLUMN].dt.to_period('M')  # Convert to Year-Month period
        # Group by Year-Month and count
        count_by_year_month = sorted_df.groupby('YearMonth').size()
        printToScreen(f"Number of records per year and month: {count_by_year_month}")

     # ******************************************************************************************
        # Create MSExcel file with the 3 dataframes
        # Create a Pandas Excel writer using XlsxWriter as the engine
        excel_file_name = output_file_selected[:-4] + "_SUMMARY.xlsx"

        # ensure that datetimes are timezone unaware before writing to Excel.
        sorted_df_noTimeZone = sorted_df
        sorted_df_noTimeZone[SERVICE_DATE_COLUMN] = pd.to_datetime(sorted_df_noTimeZone[SERVICE_DATE_COLUMN]).dt.tz_localize(None)

        with pd.ExcelWriter(excel_file_name, engine='xlsxwriter') as writer:
            # Write each dataframe to a different worksheet
            parts_replaced_df.to_excel(writer, sheet_name='Parts Replaced', index=False)
            devices_df.to_excel(writer, sheet_name='Devices', index=False)
            sorted_df_noTimeZone.to_excel(writer, sheet_name='Main', index=False)

            # Freeze the top row in each worksheet
            writer.sheets['Main'].freeze_panes(1, 0)
            writer.sheets['Parts Replaced'].freeze_panes(1, 0)
            writer.sheets['Devices'].freeze_panes(1, 0)

            # Change the tab color of the "Main" worksheet to yellow
            main_worksheet = writer.sheets['Main']
            main_worksheet.set_tab_color('yellow')
            
            # Add a new worksheet called KPIs
            workbook  = writer.book
            worksheet = workbook.add_worksheet('KPIs')
            
            # Retrieve the text from txt_edit
            kpi_text = txt_edit.get('1.0', 'end-1c')  # 'end-1c' to remove the trailing newline
            
            # Split the text into lines
            kpi_lines = kpi_text.split('\n')
            
            # Write each line into a different cell starting from A2
            for i, line in enumerate(kpi_lines):
                worksheet.write(i + 1, 0, line)  # i + 1 to start from row 2 (A2)

            # Change the tab color of the "KPIs" worksheet
            kpis_worksheet = writer.sheets['KPIs']
            kpis_worksheet.set_tab_color('blue')

        printToScreen("\n A summary MS Excel file has been created. It contains all the data and it can be used for further analysis: " + excel_file_name + "\n")    

    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        PrintException()

def do_column_overview(column_name,data_frame):
    try:

        # Check if the column exists in the DataFrame
        if column_name not in data_frame.columns:
            print(f"Column '{column_name}' does not exist in the DataFrame.")
            return


        # DETERMINE HOW MANY RECORDS DON'T HAVE VALUES
        blank_count = data_frame[column_name].isnull().sum() + (data_frame[column_name] == '').sum()
        printToScreen(f"\nNumber of blank entries in {column_name}: {blank_count} out of {data_frame.shape[0]} records, or {((blank_count / data_frame.shape[0])*100):.2f}%")   
        # DETERMINE NUMBER OF RECORDS PER VALUE TYPE
        model_counts = data_frame[column_name].value_counts()
        column_name_short = column_name.split(" - ")[1]
        printToScreen(f"\nNumber of records per {column_name_short} type: {model_counts}")




    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        PrintException()   
# CREATE GUI *************************************************************************************************

window = tk.Tk()
window.title(PROGRAM_TITLE + " - " + VERSION + '')
window.rowconfigure(0, minsize=400, weight=1)
window.columnconfigure(1, minsize=400, weight=1)
window.minsize(1200,200)

txt_edit = ScrolledText(window, width=100, height=24,wrap="word")


global fr_buttons
fr_buttons = tk.Frame(window, relief=tk.RAISED, bd=2)
btn_decode_file = tk.Button(fr_buttons, text="Select file to analyze...", command=Select_file_and_analysis)
separator = ttk.Separator(fr_buttons, orient='horizontal')

btn_select_db_file = tk.Button(fr_buttons, text="Select db file...", command=open_db_file)
#separator2 = ttk.Separator(fr_buttons, orient='horizontal')

text_box = tk.Entry(fr_buttons)

btn_decode_file.grid(row=0, column=0, sticky="ew", padx=20, pady=5)
separator.grid(row=1, column=0, sticky="ew", padx=20, pady=5)
btn_select_db_file.grid(row=2, column=0, sticky="ew", padx=20, pady=5)

fr_buttons.grid(row=0, column=0, sticky="ns")
txt_edit.grid(row=0, column=1, sticky="nsew")

status_bar = tk.Label(text='Status bar', relief=tk.RAISED)
status_bar.grid(row=3,  sticky="ew", columnspan = 2)
status_bar.config(anchor="w") # left
default_bg = status_bar['bg'] 
default_fg = status_bar['fg'] 





printToScreen(INSTRUCTIONS)
# Get the current working directory
current_directory = os.getcwd()

# Display the current working directory
current_directory_text = "Current Working Directory:" +current_directory
print(current_directory_text)
printToScreen(current_directory_text)

window.mainloop()
# END OF CREATE GUI *************************************************************************************************
