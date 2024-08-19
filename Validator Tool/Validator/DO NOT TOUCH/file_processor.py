# file_processor.py

import sys
import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import numpy as np
import glob
import re
from dateutil import parser
import warnings
import process
import File as file

def setup_module_path():
    # Get the current working directory
    current_dir = os.getcwd()

    # Get the directory path of the module
    module_path = os.path.join(current_dir, 'DO NOT TOUCH')

    # Add the module path to sys.path
    if module_path not in sys.path:
        sys.path.append(module_path)

def get_file_paths(directory_path, directory_path_B, directory_path_C, directory_path_):
    # Patterns to match both .xlsx and .xlsm files
    file_patterns = ['*.xlsx', '*.xlsm']
    
    input_path = file.get_files(directory_path_, file_patterns)
    file_paths = file.get_files(directory_path, file_patterns)
    file_paths_B = file.get_files(directory_path_B, file_patterns)
    file_paths_C = file.get_files(directory_path_C, file_patterns)

    return file_paths, file_paths_B, file_paths_C, input_path

def process_files(file_paths, file_paths_B, file_paths_C, output_directory_path, input_path):
    # Ensure at least one folder contains a file
    if not file_paths_B and not file_paths_C:
        print("Error: No files found in either 'Member Details' or 'Insurer Active Roster' directories. \n")
    
    
    if len(input_path) < 1:
        print(f"Source Folder Has no files to process.")
        print("Process Aborted")
    elif len(input_path) > 1:
        print(f"Source Folder Has More Than 2 Files: {file_paths}")
        print("Process Aborted")
    else:
        for file_path in file_paths:
            # Extract the file name without extension for output
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            
            while True:
                policy_option = input("\nPlease provide the policy option a) GMC b) GPA c) GTL: ").strip().upper()
                if policy_option in ['GMC', 'GPA', 'GTL']:
                    break
                else:
                    print("\nInvalid input. Please provide a valid policy option (GMC, GPA, GTL) \n")

            # Now you can use the `policy_option` variable
            print(f"\nSelected policy option: {policy_option} \n")

            with pd.ExcelFile(file_path) as xls:
                # Process the 'Additions' sheet
                additions_sheet = 'Additions'
                additions_df = pd.read_excel(xls, sheet_name=additions_sheet)
                additions_styled_df, additions_msg = process.process_sheet(additions_df, additions_sheet, file_paths_B, file_paths_C,
                                                                           policy_option)

                # Process the 'Deletions' sheet
                deletions_sheet = 'Deletions'
                deletions_df = pd.read_excel(xls, sheet_name=deletions_sheet)
                deletions_styled_df, deletions_msg = process.process_sheet(deletions_df, deletions_sheet, file_paths_B, file_paths_C,
                                                                           policy_option)

                # Read the 'Corrections' sheet without processing
                corrections_sheet = 'Corrections'
                corrections_df = pd.read_excel(xls, sheet_name=corrections_sheet)

                # Output the modified DataFrames to a single Excel file with multiple sheets
                output_file_path = os.path.join(output_directory_path, file_name + '.xlsx')

                with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                    if additions_msg is None:
                        additions_styled_df.to_excel(writer, sheet_name=additions_sheet, index=False)
                    else:
                        print(f"{file_name} - {additions_msg}")

                    if deletions_msg is None:
                        deletions_styled_df.to_excel(writer, sheet_name=deletions_sheet, index=False)
                    else:
                        print(f"{file_name} - {deletions_msg}")

                    corrections_df.to_excel(writer, sheet_name=corrections_sheet, index=False)
        print("\nProcess Completed")            

def main(directory_path, directory_path_B, directory_path_C, output_directory_path, directory_path_):
    # Ensures that the module's path is correctly added to sys.path. This allows the script to import necessary modules that might be in         a specific directory named DO NOT TOUCH
    setup_module_path()
    file_paths, file_paths_B, file_paths_C, input_path = get_file_paths(directory_path, directory_path_B, directory_path_C, directory_path_)
    process_files(file_paths, file_paths_B, file_paths_C, output_directory_path, input_path)
    
