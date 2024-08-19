# process.py

import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import numpy as np
import os
import sys
import glob
import re
from dateutil import parser
import warnings
import File as file
from tabulate import tabulate

# Ignoring warning message
warnings.simplefilter(action='ignore', category=pd.errors.SettingWithCopyWarning)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# Processing start for "Deletion" sheet
def process_sheet(df, sheet_name, file_paths_B, file_paths_C, policy_option):
    
    # Handling for "Deletions" sheet
    if sheet_name == "Deletions":   
        # Check if the DataFrame is empty (no data, only columns)
        if df.empty:
            return df, "Deletions: Sheet has No Data"
        
        # Remove rows with all null values
        df.dropna(how='all', inplace=True)
        
    #----------------------------------- Validator Function Begin -------------------------------------------------------------        
        
        # Check column "Emp No"
        if df['Emp No'].dtype == 'object':
            df['Emp No'] = df['Emp No'].str.replace(r'\s+', '', regex=True)
        else:
            df['Emp No'] = df['Emp No'].astype(int)

        # Convert the 'Name' column to lowercase
        df['Name'] = df['Name'].str.lower()

        # Remove specific titles from the 'Name' column
        titles_to_remove = ['mr', 'mrs', 'master', 'dr', 'miss']
        pattern = r'\b(?:' + '|'.join(re.escape(title) for title in titles_to_remove) + r')\b'
        df['Name'] = df['Name'].str.replace(pattern, '', regex=True)

        # Remove special characters from the 'Name' column
        special_chars_pattern = r'[!@#$%^&*()_+\-=\[\]{}\\|:";\'<>?,./1234567890]'
        df['Name'] = df['Name'].str.replace(special_chars_pattern, '', regex=True)
        
        # Remove extra spaces
        df['Name'] = df['Name'].str.replace(r'\s+', ' ', regex=True)

        # Strip leading/trailing whitespace and title case the 'Name' column
        df['Name'] = df['Name'].str.strip().str.title()
    
        # Clean the 'Relation' column
        df['Relation'] = df['Relation'].str.strip().str.title()
        
        # Filter the DataFrame based on the policy option
        if policy_option in ["GPA", "GTL"]:
            df = df.loc[df['Relation'].isin(['Self', 'Employee']), :]
       
        # Convert 'Date of Leaving' and 'Date of Coverage' columns to 'dd-mmm-yyyy' format
        df['Date Of Leaving'] = df['Date Of Leaving'].str.strip()  # Remove leading and trailing spaces
        df['Date Of Leaving'] = pd.to_datetime(df['Date Of Leaving'], format='mixed', dayfirst=True, errors='coerce')  # Handle mixed formats and convert to datetime
        df['Date Of Leaving'] = df['Date Of Leaving'].dt.strftime('%d/%b/%Y')  # Convert to the desired format

        
        # Check if 'Date of Coverage' is not null before calculating 'Within 45 days'
        df['Within 45 days'] = None
        not_null_mask = ~df['Date Of Leaving'].isnull()
        df.loc[not_null_mask, 'Within 45 days'] = pd.to_datetime(df.loc[not_null_mask, 'Date Of Leaving'], format='mixed', dayfirst=True, errors='coerce') >= datetime.now() - timedelta(days=42)

    #---------------------------------------- Genomo matching begins -----------------------------------------
    
        # Create a temporary column by concatenating 'Emp No' (converted to string) and 'Name'
        df['Temp_Column'] = df['Emp No'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper() +" "+ df['Name']
        # To fill any blank value in df['Temp_Column']
        df['Temp_Column'] = df['Temp_Column'].fillna('Unknown')
        
        if not file_paths_B:
            print("Process Type: Deletion: No file found for Genome \n")
        else:
            # Load the comparison file (file_paths_B) and process for comparison
            comparison_df = pd.read_excel(file_paths_B[0])

            # Initialize variables for the column names
            name_column = "Name"
            employee_id_column = "Employee ID"
            user_id_column = "User ID"
            uhid_column = "UHID"
            Relationship_column = "Relationship"
            DOB_column = "DOB"
            Gender_column = "Gender"
            Sum_Insured_column = "Sum Insured"
            Coverage_Start_Date_column = "Coverage Start Date"
            Coverage_End_Date_column = "Coverage End Date"
            Phone_column = "Phone"
            Email_column = "Email"

            # Check if the 'Name' column is present
            if name_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{name_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                name_column = input("Please enter the correct column name for 'Name': ")

            # Check if the 'Employee ID' column is present
            if employee_id_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{employee_id_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                employee_id_column = input("Please enter the correct column name for 'Employee ID': ")
                
             # Check if the "User ID" column is present
            if user_id_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{user_id_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                user_id_column = input("Please enter the correct column user id for 'User ID': ")
                
             # Check if the "Coverage Start Date" column is present
            if Coverage_Start_Date_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{Coverage_Start_Date_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                Coverage_Start_Date_column = input("Please enter the correct column Coverage Start Date for 'Coverage Start Date': ")
                
             # Check if the "Relationship" column is present
            if Relationship_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{Relationship_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                Relationship_column = input("Please enter the correct column Relationship for 'Relationship': ")
 
             # Check if the "DOB" column is present
            if DOB_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{DOB_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                DOB_column = input("Please enter the correct column DOB for 'DOB': ")
                
             # Check if the "Gender" column is present
            if Gender_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{Gender_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                Gender_column = input("Please enter the correct column Gender for 'Gender': ")                
                
             # Check if the "Sum Insured" column is present
            if Sum_Insured_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{Sum_Insured_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                Sum_Insured_column = input("Please enter the correct column Sum Insured for 'Sum Insured': ")
                
             # Check if the "Coverage Start Date" column is present
            if Coverage_Start_Date_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{Coverage_Start_Date_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                Coverage_Start_Date_column = input("Please enter the correct column Coverage Start Date for 'Coverage Start Date': ")                
                
             # Check if the "Coverage End Date" column is present
            if Coverage_End_Date_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{Coverage_End_Date_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                Coverage_End_Date_column = input("Please enter the correct column Coverage End Date for 'Coverage End Date': ")               
                
             # Check if the "Phone" column is present
            if Phone_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{Phone_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                Phone_column = input("Please enter the correct column Phone for 'Phone': ")                

             # Check if the "Email" column is present
            if Email_column not in comparison_df.columns:
                print(f"The file does not contain the required column: '{Email_column}' \n")
                print(f"Columns present in the file: {comparison_df.columns.tolist()} \n")
                # Ask for input for the correct column name
                Email_column = input("Please enter the correct column Email for 'Email': ")                
                
                
            
            # Remove rows with all null values
            comparison_df.dropna(how='all', inplace=True)

            # Convert the 'Name' column to lowercase
            comparison_df[name_column] = comparison_df[name_column].str.lower()

            # Remove specific titles from the 'Name' column
            comparison_df[name_column] = comparison_df[name_column].str.replace(pattern, '', regex=True)

            # Remove special characters from the 'Name' column
            comparison_df[name_column] = comparison_df[name_column].str.replace(special_chars_pattern, '', regex=True)
            
            # Remove extra spaces
            comparison_df[name_column] = comparison_df[name_column].str.replace(r'\s+', ' ', regex=True)

            # Strip leading/trailing whitespace and title case the 'Name' column
            comparison_df[name_column] = comparison_df[name_column].str.strip().str.title()

            # Convert 'Employee ID' column to string, remove any whitespace, and remove .0
            comparison_df[employee_id_column] = comparison_df[employee_id_column].astype(str).str.replace(r'\.0$', '', regex=True)

            # Convert 'Name' column to string
            comparison_df[name_column] = comparison_df[name_column].astype(str)
        
            # Create a temporary column by concatenating 'Employee ID' and 'Name'
            comparison_df['Temp_Column'] = comparison_df[employee_id_column].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper() + " " + comparison_df[name_column]
        
            # Getting User ID and UHID column from Genome data
            # Columns
            active_column = 'Active'
            Relationship_column = "Relationship"
            DOB_column = "DOB"
            Gender_column = "Gender"
            Sum_Insured_column = "Sum Insured"
            Coverage_Start_Date_column = "Coverage Start Date"
            Coverage_End_Date_column = "Coverage End Date"
            Phone_column = "Phone"
            Email_column = "Email"
            

            # Check if the columns exist in the comparison_df
            uhid_exists = uhid_column in comparison_df.columns
            active_exists = active_column in comparison_df.columns
            Relationship_exists = Relationship_column in comparison_df.columns
            DOB_exists = DOB_column in comparison_df.columns
            Gender_exists = Gender_column in comparison_df.columns
            Sum_Insured_exists = Sum_Insured_column in comparison_df.columns
            Coverage_Start_Date_exists = Coverage_Start_Date_column in comparison_df.columns
            Coverage_End_Date_exists = Coverage_End_Date_column in comparison_df.columns
            Phone_exists = Phone_column in comparison_df.columns
            Email_exists = Email_column in comparison_df.columns

            def match_names(df_temp, comp_temp, comp_user_id, comp_uhid=None, comp_active=None, comp_Relationship=None, comp_DOB=None, comp_Gender=None, comp_Sum_Insured=None, comp_Coverage_Start_Date=None, comp_Coverage_End_Date=None, comp_Phone=None, comp_Email=None):
                df_parts = df_temp.split()
                comp_parts = comp_temp.split()

                # Check if Employee ID matches
                if df_parts[:2] != comp_parts[:2]:
                    return False, None, None, None, None, None, None, None, None, None, None, None

                # Compare the names
                df_name_parts = set(df_parts[2:])
                comp_name_parts = set(comp_parts[2:])

                # Finding common name parts
                common_parts = df_name_parts.intersection(comp_name_parts)

                # We could also consider common_parts >= 1 if we want to be less strict
                if len(common_parts) >= 1:
                    return True, comp_user_id, comp_uhid, comp_active, comp_Relationship, comp_DOB, comp_Gender, comp_Sum_Insured, comp_Coverage_Start_Date, comp_Coverage_End_Date, comp_Phone, comp_Email
                else:
                    return False, None, None, None, None, None, None, None, None, None, None, None

            def get_matchs(df_temp, comparison_df):
                for i, row in comparison_df.iterrows():
                    match, user_id, uhid, active, Relationship, DOB, Gender, Sum_Insured, Coverage_Start_Date, Coverage_End_Date, Phone, Email  = match_names(
                        df_temp,
                        row['Temp_Column'],
                        row[user_id_column],
                        row[uhid_column] if uhid_exists else None,
                        row[active_column] if active_exists else None,
                        row[Relationship_column] if Relationship_exists else None,
                        row[DOB_column] if DOB_exists else None,
                        row[Gender_column] if Gender_exists else None,
                        row[Sum_Insured_column] if Sum_Insured_exists else None,
                        row[Coverage_Start_Date_column] if Coverage_Start_Date_exists else None,
                        row[Coverage_End_Date_column] if Coverage_End_Date_exists else None,
                        row[Phone_column] if Phone_exists else None,
                        row[Email_column] if Email_exists else None
                    )
                    if match:
                        return row['Temp_Column']
                return None

            # Apply the matching function
            df['Match Found Genome'] = df['Temp_Column'].apply(
                lambda x: any(
                    match_names(
                        x,
                        row['Temp_Column'],
                        row[user_id_column],
                        row[uhid_column] if uhid_exists else None,
                        row[active_column] if active_exists else None,
                        row[Relationship_column] if Relationship_exists else None,
                        row[DOB_column] if DOB_exists else None,
                        row[Gender_column] if Gender_exists else None,
                        row[Sum_Insured_column] if Sum_Insured_exists else None,
                        row[Coverage_Start_Date_column] if Coverage_Start_Date_exists else None,
                        row[Coverage_End_Date_column] if Coverage_End_Date_exists else None,
                        row[Phone_column] if Phone_exists else None,
                        row[Email_column] if Email_exists else None
                    )[0] for _, row in comparison_df.iterrows()
                )
            )

            # Extra added Start
            df['Found on Genome'] = df['Temp_Column'].apply(
                lambda x: get_matchs(x, comparison_df) if any(
                    match_names(
                        x,
                        row['Temp_Column'],
                        row[user_id_column],
                        row[uhid_column] if uhid_exists else None,
                        row[active_column] if active_exists else None,
                        row[Relationship_column] if Relationship_exists else None,
                        row[DOB_column] if DOB_exists else None,
                        row[Gender_column] if Gender_exists else None,
                        row[Sum_Insured_column] if Sum_Insured_exists else None,
                        row[Coverage_Start_Date_column] if Coverage_Start_Date_exists else None,
                        row[Coverage_End_Date_column] if Coverage_End_Date_exists else None,
                        row[Phone_column] if Phone_exists else None,
                        row[Email_column] if Email_exists else None
                    )[0] for _, row in comparison_df.iterrows()
                ) else None
            )

            # Apply the matching function and retrieve User ID, UHID, and Active status
            def get_matching_info(temp_value):
                for i, row in comparison_df.iterrows():
                    match, user_id, uhid, active, Relationship, DOB, Gender, Sum_Insured, Coverage_Start_Date, Coverage_End_Date, Phone, Email = match_names(
                        temp_value,
                        row['Temp_Column'],
                        row[user_id_column],
                        row[uhid_column] if uhid_exists else None,
                        row[active_column] if active_exists else None,
                        row[Relationship_column] if Relationship_exists else None,
                        row[DOB_column] if DOB_exists else None,
                        row[Gender_column] if Gender_exists else None,
                        row[Sum_Insured_column] if Sum_Insured_exists else None,
                        row[Coverage_Start_Date_column] if Coverage_Start_Date_exists else None,
                        row[Coverage_End_Date_column] if Coverage_End_Date_exists else None,
                        row[Phone_column] if Phone_exists else None,
                        row[Email_column] if Email_exists else None
                    )
                    if match:
                        return user_id, uhid, active, Relationship, DOB, Gender, Sum_Insured, Coverage_Start_Date, Coverage_End_Date, Phone, Email
                return None, None, None, None, None, None, None, None, None, None, None

            df['User ID'], df['UHID'], df['Active'], df['Relationship'], df['DOB'], df['Gender'], df['Sum Insured'], df['Coverage Start Date'], df['Coverage End Date'], df['Phone'], df['Email'] = zip(*df['Temp_Column'].map(get_matching_info))

 
                
    #-------------------------------------------------------------------------------------------------------------------------
    
        # End of Delete Process
        return df, None
    
    #----------------------------------- Sheet Addition Function Begin ------------------------------------------------------
    
    # Check if the DataFrame is empty (no data, only columns)
    if df.empty:
        return df, "Additions: Sheet has No Data"
    
    # Remove rows with all null values
    df.dropna(how='all', inplace=True)
    
    #----------------------------------- Validator Function Begin -------------------------------------------------------------
    
    # Check column "Emp No"
    if df['Emp No'].dtype == 'object':
        df['Emp No'] = df['Emp No'].str.replace(r'\s+', '', regex=True)
    else:
        df['Emp No'] = df['Emp No'].astype(int)
    
    # Convert the 'Name' column to lowercase
    df['Name'] = df['Name'].str.lower()

    # Remove specific titles from the 'Name' column
    titles_to_remove = ['mr', 'mrs', 'master', 'dr', 'miss']
    pattern = r'\b(?:' + '|'.join(re.escape(title) for title in titles_to_remove) + r')\b'
    df['Name'] = df['Name'].str.replace(pattern, '', regex=True)

    # Remove special characters from the 'Name' column
    special_chars_pattern = r'[!@#$%^&*()_+\-=\[\]{}\\|:";\'<>?,./1234567890]'
    df['Name'] = df['Name'].str.replace(special_chars_pattern, '', regex=True)
    
   # Remove extra spaces
    df['Name'] = df['Name'].str.replace(r'\s+', ' ', regex=True)

    # Strip leading/trailing whitespace and title case the 'Name' column
    df['Name'] = df['Name'].str.strip().str.title()

    # Clean the 'Relation' column
    df['Relation'] = df['Relation'].str.strip().str.title()
    
    # Filter the DataFrame based on the policy option
    if policy_option in ["GPA", "GTL"]:
        df = df.loc[df['Relation'].isin(['Self', 'Employee']), :]
        
    # Convert 'Relation' column values
    def convert_gender(gender):
        if gender in ['Female', 'F', 'female', 'FEMALE']:
            return 'Female'
        elif gender in ['Male', 'M', 'MALE', 'male']:
            return 'Male'
        else:
            return gender
    
    # Clean the 'Gender' column
    df['Gender'] = df['Gender'].str.strip().str.title().apply(convert_gender)

    # Convert 'Relation' column values
    def convert_relation(relation):
        if relation == 'Employee':
            return 'Self'
        elif relation == 'Wife':
            return 'Spouse'
        elif relation in ['Daughter', 'Son']:
            return 'Child'
        elif relation in ['Father', 'Mother']:
            return 'Parent'
        elif relation in ['Father In Law', 'Mother In Law']:
            return 'Parent-in-law'
        else:
            return relation

    df['Relation Type'] = df['Relation'].apply(convert_relation)
    
    # Reorder columns to have 'Relation Type' next to 'Relation'
    cols = df.columns.tolist()
    relation_index = cols.index('Relation')
    cols.insert(relation_index + 1, cols.pop(cols.index('Relation Type')))
    df = df[cols]

    # Convert 'Date of Birth' and 'Date of Coverage' columns to 'dd-mmm-yyyy' format
    def convert_to_dd_mmm_yyyy(date_str):
        if pd.isnull(date_str):
            return None
        try:
            # Parse the date using dateutil.parser
            date = parser.parse(date_str)
            # Format the date to 'DD-MMM-YYYY'
            return date.strftime('%d/%b/%Y')
        except Exception as e:
            print(f"Error parsing date: {date_str}. Exception: {e}")
            return None  # or handle the exception as needed

    # Apply the function to 'Date of Birth' and 'Date of Coverage' columns
    df['Date Of Birth'] = df['Date Of Birth'].str.strip()  # Remove leading and trailing spaces
    df['Date Of Birth'] = df['Date Of Birth'].apply(convert_to_dd_mmm_yyyy)

    
    if 'Date Of Joining' in df.columns:
        if not df['Date Of Joining'].isna().all():
            df['Date Of Joining'] = df['Date Of Joining'].str.strip()  # Remove leading and trailing spaces
            df['Date Of Joining'] = df['Date Of Joining'].apply(convert_to_dd_mmm_yyyy)
            df['Date Of Joining'] = df.groupby('Emp No')['Date Of Joining'].ffill()
        
    if 'Date Of Coverage' in df.columns:
        if not df['Date Of Coverage'].isna().all():
            df['Date Of Coverage'] = df['Date Of Coverage'].str.strip()  # Remove leading and trailing spaces
            df['Date Of Coverage'] = df['Date Of Coverage'].apply(convert_to_dd_mmm_yyyy)
            df['Date Of Coverage'] = df.groupby('Emp No')['Date Of Coverage'].ffill()

    # Check if both 'Date Of Joining' and 'Date Of Coverage' are not present
    if 'Date Of Joining' not in df.columns and 'Date Of Coverage' not in df.columns:
        print("Columns are empty. Aborting execution.")
        

    # Delete cells from 'Email ID' and 'Mobile Number' columns based on 'Relation'
    df.loc[df['Relation Type'] != 'Self', ['Email Id', 'Mobile Number']] = np.nan

    # Create 'Age' column from 'Date of Birth' column
    df['Age'] = pd.to_datetime(df['Date Of Birth'], format='mixed', dayfirst=True, errors='coerce')
    df['Age'] = ((datetime.now() - df['Age']).dt.days // 365).fillna(0).astype(int)

    # Check if 'Date of Coverage' is not null before calculating 'Within 45 days'
    if 'Date Of Joining' in df.columns:
        if not df['Date Of Joining'].isna().all():
            df['Within 45 days'] = None
            not_null_mask = ~df['Date Of Joining'].isnull()
            df.loc[not_null_mask, 'Within 45 days'] = pd.to_datetime(df.loc[not_null_mask, 'Date Of Joining'], format='mixed', dayfirst=True, errors='coerce') >= datetime.now() - timedelta(days=42)
    else:
        df['Within 45 days'] = None
        not_null_mask = ~df['Date Of Coverage'].isnull()
        df.loc[not_null_mask, 'Within 45 days'] = pd.to_datetime(df.loc[not_null_mask, 'Date Of Coverage'], format='mixed', dayfirst=True, errors='coerce') >= datetime.now() - timedelta(days=42)
    
    # Clean the 'Email ID' column
    if 'Email Id' in df.columns:
        if not df['Email Id'].isna().all():
            df['Email Id'] = df['Email Id'].str.lower()

    # New logic for coloring rows based on Emp No and Relation
    grouped_df = df.groupby(['Emp No', 'Relation']).size().reset_index(name='count')
    emp_no_with_no_self = grouped_df.groupby('Emp No').filter(lambda x: all(rel not in ['Self', 'Employee'] for rel in x['Relation']))

    # Create a set of Emp Nos that meet the criteria
    emp_no_set = set(emp_no_with_no_self['Emp No'])
    
    # Reorder columns to have 'Gender' next to 'Relation'
    cols = df.columns.tolist()
    relation_type_index = cols.index('Relation Type')
    cols.insert(relation_type_index + 1, cols.pop(cols.index('Gender')))
    df = df[cols]

    #----------------------------------- Error Message Function Begin -------------------------------------------------------------
    
    # Step 2: Create 'Missing Error' column and initialize with empty strings
    df['Missing Error'] = ''

    # a) Check for duplicate names and add error message
    duplicate_names = df['Name'].duplicated(keep=False)
    df.loc[duplicate_names, 'Missing Error'] += 'Duplicate Name Found; '

    # b) Check if 'Date of Birth' is missing for given 'Emp No' and 'Name'
    missing_dob = df['Date Of Birth'].isnull() & df['Emp No'].notnull() & df['Name'].notnull()
    df.loc[missing_dob, 'Missing Error'] += 'Birth date Missing; '

    # c) Check if 'Name' is missing for given 'Emp No' and 'Date of Birth'
    missing_name = df['Name'].isnull() & df['Emp No'].notnull() & df['Date Of Birth'].notnull()
    df.loc[missing_name, 'Missing Error'] += 'Name is Missing; '

    # d) Check if 'Date of Coverage' is missing for 'Employee' or 'Self' in 'Relation'
    if 'Date Of Coverage' in df.columns:
        if not df['Date Of Coverage'].isna().all():
            missing_coverage = df['Date Of Coverage'].isnull() & ((df['Relation'] == 'Employee') | (df['Relation Type'] == 'Self'))
            df.loc[missing_coverage, 'Missing Error'] += 'Date of Coverage is missing; '

    # e) Check if 'Relation' is missing for given 'Emp No' and 'Name'
    missing_relation = df['Relation'].isnull() & df['Emp No'].notnull() & df['Name'].notnull()
    df.loc[missing_relation, 'Missing Error'] += 'Relation is Missing; '

    # f) Check if Email ID does not contain '@'
    if 'Email Id' in df.columns:
        if not df['Email Id'].isna().all():
            invalid_email = df['Email Id'].notnull() & ~df['Email Id'].str.contains('@', na=False)
            df.loc[invalid_email, 'Missing Error'] += 'Email Is not Valid; '
            
            # Check for missing mobile numbers for 'Employee' and 'Self'
            email_id_for_emp_no_self = df['Relation'].isin(['Employee', 'Self']) & df['Email Id'].isnull()
            df.loc[email_id_for_emp_no_self, 'Missing Error'] += 'Email ID is Missing; '

            # Duplicate Email
            emp_no_self = df[df['Relation'].isin(['Employee', 'Self'])]
            emp_no_self_with_email = emp_no_self[emp_no_self['Email Id'].notnull()]
            duplicate_emp_no_self = emp_no_self_with_email.duplicated(subset=['Email Id'], keep=False)
            df.loc[emp_no_self_with_email.index[duplicate_emp_no_self], 'Missing Error'] += 'Duplicate Email ID Found; '

    # g) Check if 'Emp No' is the same when 'Relation' is 'Self' and add error message
    emp_no_self = df[df['Relation'].isin(['Employee', 'Self'])]
    duplicate_emp_no_self = emp_no_self.duplicated(subset=['Emp No'], keep=False)
    df.loc[emp_no_self.index[duplicate_emp_no_self], 'Missing Error'] += 'Duplicate Employee Number Found; '
    
    # h) Check the age difference between an employee and their parent
    parents = df[df['Relation'].isin(['Parent', 'Father', 'Mother'])]
    for idx, row in parents.iterrows():
        emp_no = row['Emp No']
        parent_age = row['Age']
    
        employee = df[(df['Emp No'] == emp_no) & (df['Relation'].isin(['Employee', 'Self']))]
        if not employee.empty:
            employee_age = employee['Age'].values[0]
            if (parent_age - employee_age) < 16:
                df.loc[idx, 'Missing Error'] += 'Age difference between employee and parent is less than 16 years; '
                
                
    # i) Check if the gender is the same for 'Self' and 'Spouse' for the same 'Emp No'
    emp_no_gender = df[df['Relation'].isin(['Self', 'Spouse'])][['Emp No', 'Relation', 'Gender']]

    for emp_no in emp_no_gender['Emp No'].unique():
        self_gender = emp_no_gender[(emp_no_gender['Emp No'] == emp_no) & (emp_no_gender['Relation'] == 'Self')]['Gender']
        spouse_gender = emp_no_gender[(emp_no_gender['Emp No'] == emp_no) & (emp_no_gender['Relation'] == 'Spouse')]['Gender']
    
        if not self_gender.empty and not spouse_gender.empty:
            if self_gender.iloc[0] == spouse_gender.iloc[0]:
                self_index = emp_no_gender[(emp_no_gender['Emp No'] == emp_no) & (emp_no_gender['Relation'] == 'Self')].index
                spouse_index = emp_no_gender[(emp_no_gender['Emp No'] == emp_no) & (emp_no_gender['Relation'] == 'Spouse')].index
                df.loc[self_index, 'Missing Error'] += 'Self and Spouse Gender are the same; '
                df.loc[spouse_index, 'Missing Error'] += 'Self and Spouse Gender are the same; '


    #------------------------------------------ Validation Function Begin -------------------------------------------------------

    # g) Check if Mobile Number is not valid
    def preprocess_mobile_number(number):
        if pd.isna(number):
            return number
        str_number = str(number).strip()  # Convert to string and strip any surrounding whitespace
        if str_number.endswith('.0'):
            str_number = str_number[:-2]  # Remove trailing '.0' if present
        elif str_number.startswith('+91'):
            str_number = str_number[3:]
            
        # Remove any internal spaces and return the number if it only contains digits
        str_number = str_number.replace(' ', '')  # Remove any spaces within the number

        return str_number if str_number.isdigit() else number  # Return as string if it's a valid number

    # Apply preprocessing to the 'Mobile Number' column
    df['Mobile Number'] = df['Mobile Number'].apply(preprocess_mobile_number)

    def is_valid_mobile(number):
        if pd.isna(number):
            return True  # Treat missing numbers as valid
        number_str = str(number)
        return len(number_str) == 10 and number_str.isdigit()

    # Convert valid mobile numbers to integers and validate
    def convert_to_integer_if_valid(number):
        if pd.isna(number):
            return number  # Return NaN as is
        if is_valid_mobile(number):
            return int(number)  # Convert valid number to integer
        return number  # Keep invalid number as is

    # Check each mobile number and print if invalid
    for index, number in enumerate(df['Mobile Number']):
        if not is_valid_mobile(number):
            print(f"Invalid mobile number at index {index}: {number}")

    # Mark invalid mobile numbers in the 'Missing Error' column
    invalid_mobile = df['Mobile Number'].notnull() & ~df['Mobile Number'].apply(is_valid_mobile)
    df.loc[invalid_mobile, 'Missing Error'] += 'Mobile Number is Not valid; '
    
    # Check for missing mobile numbers for 'Employee' and 'Self'
    missing_mobile_for_emp_no_self = df['Relation'].isin(['Employee', 'Self']) & df['Mobile Number'].isnull()
    df.loc[missing_mobile_for_emp_no_self, 'Missing Error'] += 'Mobile Number is Missing; '

    
    emp_no_self = df[df['Relation'].isin(['Employee', 'Self'])]
    emp_no_self_with_mobile = emp_no_self[emp_no_self['Mobile Number'].notnull()]
    duplicate_emp_no_self = emp_no_self_with_mobile.duplicated(subset=['Mobile Number'], keep=False)
    df.loc[emp_no_self_with_mobile.index[duplicate_emp_no_self], 'Missing Error'] += 'Duplicate Mobile Number Found; '

    # Convert valid mobile numbers to integers
    df['Mobile Number'] = df['Mobile Number'].apply(convert_to_integer_if_valid)    

    # Remove trailing semicolon and space from 'Missing Error' column
    df['Missing Error'] = df['Missing Error'].str.rstrip('; ')
    
#----------------------------------------Genomo file matching function begin---------------------------------------------

    # Create a temporary column by concatenating 'Emp No' (converted to string) and 'Name'
    df['Temp_Column'] = df['Emp No'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper() + " " + df['Date Of Birth'] +" "+ df['Name']

    # To fill any blank value in df['Temp_Column']
    df['Temp_Column'] = df['Temp_Column'].fillna('Unknown')

    if not file_paths_B:
        print("Process Type: Addition: No file found for Genemo")
    else:
         # Load the comparison file (file_paths_B) and process for comparison
        comparison_df = pd.read_excel(file_paths_B[0])

        # Initialize variables for the column names
        name_column = "Name"
        employee_id_column = "Employee ID"
        date_of_birth_column = "DOB"
        active_column = "Active"
        user_id_column = "User ID"


        # Check if the 'Name' column is present
        if name_column not in comparison_df.columns:
            print(f"The file does not contain the required column: '{name_column}'\n")
            print(f"Columns present in the file: {comparison_df.columns.tolist()}\n")
            # Ask for input for the correct column name
            name_column = input("Please enter the correct column name for 'Name': ")

        # Check if the 'Employee ID' column is present
        if employee_id_column not in comparison_df.columns:
            print(f"The file does not contain the required column: '{employee_id_column}'\n")
            print(f"Columns present in the file: {comparison_df.columns.tolist()}\n")
            # Ask for input for the correct column name
            employee_id_column = input("Please enter the correct column name for 'Employee ID': ")
            
        # Check if the 'Birth Date' column is present
        if date_of_birth_column not in comparison_df.columns:
            print(f"The file does not contain the required column: '{date_of_birth_column}'\n")
            print(f"Columns present in the file: {comparison_df.columns.tolist()}\n")
            # Ask for input for the correct column name
            date_of_birth_column = input("Please enter the correct column name for 'Date Of Birth': ")
            
        # Check if the 'Active' column is present
        if active_column not in comparison_df.columns:
            print(f"The file does not contain the required column: '{active_column}'\n")
            print(f"Columns present in the file: {comparison_df.columns.tolist()}\n")
            # Ask for input for the correct column name
            active_column = input("Please enter the correct column name for 'Active': ")
            
            
        # Check if the 'Active' column is present
        if user_id_column not in comparison_df.columns:
            print(f"The file does not contain the required column: '{user_id_column}'\n")
            print(f"Columns present in the file: {comparison_df.columns.tolist()}\n")
            # Ask for input for the correct column name
            user_id_column = input("Please enter the correct column name for 'Active': ")
            
        # Remove rows with all null values
        comparison_df.dropna(how='all', inplace=True)

        # Convert the 'Name' column to lowercase
        comparison_df[name_column] = comparison_df[name_column].str.lower()

        # Remove specific titles from the 'Name' column
        comparison_df[name_column] = comparison_df[name_column].str.replace(pattern, '', regex=True)

        # Remove special characters from the 'Name' column
        comparison_df[name_column] = comparison_df[name_column].str.replace(special_chars_pattern, '', regex=True)
        
        # Remove extra spaces
        comparison_df[name_column] = comparison_df[name_column].str.replace(r'\s+', ' ', regex=True)

        # Strip leading/trailing whitespace and title case the 'Name' column
        comparison_df[name_column] = comparison_df[name_column].str.strip().str.title()

        # Convert 'Employee ID' column to string, remove any whitespace, and remove .0
        comparison_df[employee_id_column] = comparison_df[employee_id_column].astype(str).str.replace(r'\.0$', '', regex=True)

        # Convert 'Name' column to string
        comparison_df[name_column] = comparison_df[name_column].astype(str)
        
        
        # Convert 'Date of Birth' columns to 'dd/mmm/yyyy' format
        def convert_to_dd_mmm_yyyy(date_str):
            if pd.isnull(date_str):
                return None
            try:
                # Parse the date using dateutil.parser
                date = parser.parse(date_str)
                # Format the date to 'DD-MMM-YYYY'
                return date.strftime('%d/%b/%Y')
            except Exception as e:
                print(f"Error parsing date: {date_str}. Exception: {e}")
                return None  # or handle the exception as needed

        # Apply the function to 'Date of Birth' and 'Date of Coverage' columns
        comparison_df[date_of_birth_column] = comparison_df[date_of_birth_column].str.strip()  # Remove leading and trailing spaces
        comparison_df[date_of_birth_column] = comparison_df[date_of_birth_column].apply(convert_to_dd_mmm_yyyy)
        
        # Create a temporary column by concatenating 'Employee ID' and 'Name'
        comparison_df['Temp_Column'] = comparison_df[employee_id_column].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()  + " " + comparison_df[date_of_birth_column] + " " + comparison_df[name_column]
        
        # Function to match names with configurable common part threshold
        def match_names(df_temp, comp_temp, min_common_parts=1):
            df_parts = [part.strip() for part in df_temp.split()]
            comp_parts = [part.strip() for part in comp_temp.split()]

            if df_parts[:2] != comp_parts[:2]:
                return False

            df_name_parts = set(df_parts[2:])
            comp_name_parts = set(comp_parts[2:])

            common_parts = df_name_parts.intersection(comp_name_parts)

            return len(common_parts) >= min_common_parts

        # Apply the matching function
        df['Match Found Genome'] = df['Temp_Column'].apply(
            lambda x: any(match_names(x, y) for y in comparison_df['Temp_Column'])
        )

        # Map Active status based on matching criteria
        def get_active_status(row):
            matches = comparison_df.loc[comparison_df['Temp_Column'].apply(lambda y: match_names(row['Temp_Column'], y))]
            if not matches.empty:
                if active_column in matches.columns:
                    return matches[active_column].values[0]  # Access the active_column safely
                else:
                    return None  # Or handle it in some other way if the column doesn't exist
            else:
                return None  # Or handle the case where no match is found

        df['Active'] = df.apply(get_active_status, axis=1)
        
        # Map Active status based on matching criteria
        def get_user_id_status(row):
            matches = comparison_df.loc[comparison_df['Temp_Column'].apply(lambda y: match_names(row['Temp_Column'], y))]
            if not matches.empty:
                if user_id_column in matches.columns:
                    return matches[user_id_column].values[0]  # Access the user_id_column safely
                else:
                    return None  # Or handle it in some other way if the column doesn't exist
            else:
                return None  # Or handle the case where no match is found

        df['User ID'] = df.apply(get_user_id_status, axis=1)
        
        def get_match(df_temp, comparison_df):
            for comp_temp in comparison_df['Temp_Column']:
                if match_names(df_temp, comp_temp):
                    return comp_temp
            return None
        
        df['Found on Genome'] = df['Temp_Column'].apply(
            lambda x: get_match(x, comparison_df) if any(match_names(x, y) for y in comparison_df['Temp_Column']) else None
        )
        
#-------------------------------------------------------------------------------------------        
        # Function to match names with configurable common part threshold
        def match_name(df_temp, comp_temp, min_common_parts=3):
            df_parts = [part.strip() for part in df_temp.split()]
            comp_parts = [part.strip() for part in comp_temp.split()]

            if df_parts[0] != comp_parts[0]:
                return False

            df_name_parts = set(df_parts[1:])
            comp_name_parts = set(comp_parts[1:])

            common_parts = df_name_parts.intersection(comp_name_parts)

            return len(common_parts) >= min_common_parts

        # Apply the matching function
        df['Match Found Genome1'] = df['Temp_Column'].apply(
            lambda x: any(match_name(x, y) for y in comparison_df['Temp_Column'])
        )

        # Map Active status based on matching criteria
        def get_active_status(row):
            matches = comparison_df.loc[comparison_df['Temp_Column'].apply(lambda y: match_name(row['Temp_Column'], y))]
            if not matches.empty:
                if active_column in matches.columns:
                    return matches[active_column].values[0]  # Access the active_column safely
                else:
                    return None  # Or handle it in some other way if the column doesn't exist
            else:
                return None  # Or handle the case where no match is found
            
            
        df['Active1'] = df.apply(get_active_status, axis=1)
        
        # Map Active status based on matching criteria
        def get_userid_status(row):
            matches = comparison_df.loc[comparison_df['Temp_Column'].apply(lambda y: match_name(row['Temp_Column'], y))]
            if not matches.empty:
                if user_id_column in matches.columns:
                    return matches[user_id_column].values[0]  # Access the user_id_column safely
                else:
                    return None  # Or handle it in some other way if the column doesn't exist
            else:
                return None  # Or handle the case where no match is found
            
            
        df['User ID1'] = df.apply(get_userid_status, axis=1)
        
        def get_match(df_temp, comparison_df):
            for comp_temp in comparison_df['Temp_Column']:
                if match_name(df_temp, comp_temp):
                    return comp_temp
            return None
        
        df['Found on Genome1'] = df['Temp_Column'].apply(
            lambda x: get_match(x, comparison_df) if any(match_name(x, y) for y in comparison_df['Temp_Column']) else None
        )
  
    #-----------------------------------Active Roster file matching function begin----------------------------------------
 
    def load_comparison_sheet(file_path):
        try:
            comparison_sheets = pd.read_excel(file_path, sheet_name=None)
            # Function to add sheet name to the list.
            for sheet_name in [" "]:
                if sheet_name in comparison_sheets and not comparison_sheets[sheet_name].dropna(how='all').empty:
                    return comparison_sheets[sheet_name], comparison_sheets
            return None, comparison_sheets
        except Exception as e:
            print(f"Error loading comparison sheet: {e}")
            return None, None

    
    def get_valid_sheet(comparison_sheets):
        while True:
            print(f"Available sheets: {list(comparison_sheets.keys())}")
            user_sheet = input("Please input the sheet name to use: \n").strip()
            if user_sheet in comparison_sheets and not comparison_sheets[user_sheet].dropna(how='all').empty:
                return comparison_sheets[user_sheet]
            else:
                print("Provided sheet name is either not present or empty. \n")
                user_choice = input("Are you providing another sheet name (yes) or exiting (no)?").lower()
                if user_choice == 'no':
                    return None

    def validate_columns(comparison_df):
        required_columns = ["Name", "Emp No", "DOB", "Coverage Status"]
        missing_columns = [col for col in required_columns if col not in comparison_df.columns]

        # Initialize variables to None
        name_column = None
        emp_no_column = None
        DOB_column = None
        active_column = None

        if missing_columns:
            # Printing DataFrame in Tabular format
            pd.set_option("display.max_columns", None)
            display(comparison_df.head())
            
            # Print missing and available columns using tabulate in JIRA format
            print(f"\nColumns missing: {missing_columns}.\n\nAvailable columns:\n")
            print(tabulate([comparison_df.columns], headers='firstrow', tablefmt='jira'))
            print("\n")
            
            for missing in missing_columns:
                while True:
                    if missing == "Name":
                        name_column = input("Please input the column name for 'Name': ").strip()
                        if name_column in comparison_df.columns:
                            break
                    elif missing == "Emp No":
                        emp_no_column = input("Please input the column name for 'Emp No': ").strip()
                        if emp_no_column in comparison_df.columns:
                            break
                    elif missing == "DOB":
                        DOB_column = input("Please input the column name for 'Date Of Birth': ").strip()
                        if DOB_column in comparison_df.columns:
                            break
                    elif missing == "Coverage Status":
                        active_column = input("Please input the column name for 'Coverage Status': ").strip()
                        if active_column in comparison_df.columns:
                            break
                    print("Invalid column name. Please try again.")
        else:
            name_column = "Name"
            emp_no_column = "Emp No"
            DOB_column = "DOB"
            active_column = "Coverage Status"
    
        # Handle cases where user didn't need to input because they were already there
        if name_column is None:
            name_column = "Name"
        if emp_no_column is None:
            emp_no_column = "Emp No"
        if DOB_column is None:
            DOB_column = "DOB"
        if active_column is None:
            active_column = "Coverage Status"

        # Final check to verify all columns are correctly assigned and exist in DataFrame
        if all(col in comparison_df.columns for col in [name_column, emp_no_column, DOB_column, active_column]):
            return name_column, emp_no_column, DOB_column, active_column
        else:
            print("One or more provided column names do not exist in the DataFrame after user input. Exiting.")
            return None, None, None, None
    


    def clean_column(comparison_df, column_name):
        return comparison_df[column_name].str.strip().str.replace('\s+', ' ', regex=True).str.title()

    def convert_to_dd_mmm_yyyy(date_str):
        if pd.isnull(date_str):
            return None
        try:
            date = parser.parse(date_str)
            return date.strftime('%d-%b-%Y')
        except Exception as e:
            print(f"Error parsing date: {date_str}. Exception: {e}")
            return None

    def process_comparison_df(comparison_df, name_column, emp_no_column, DOB_column, active_column):
            
        if name_column in comparison_df.columns:
            comparison_df[name_column] = (
            comparison_df[name_column]
            .astype(str)
            .str.lower()
            .str.replace(r'[^\w\s]', '', regex=True)  # Remove special characters
            .str.replace(r'\s+', ' ', regex=True)  # Replace multiple spaces with a single space
            .str.strip()  # Strip leading/trailing whitespace
            .str.title()  # Title case the names
        )


        if emp_no_column in comparison_df.columns:
            comparison_df[emp_no_column] = comparison_df[emp_no_column].astype(str).str.replace(r'\.0$', '', regex=True)

        if DOB_column in comparison_df.columns:
            comparison_df[DOB_column] = comparison_df[DOB_column].astype(str).str.strip()
            comparison_df[DOB_column] = comparison_df[DOB_column].apply(convert_to_dd_mmm_yyyy)

        if emp_no_column in comparison_df.columns and DOB_column in comparison_df.columns and name_column in comparison_df.columns:
            comparison_df['Temp_Column'] = (
                comparison_df[emp_no_column].astype(str)
                .str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper() + " " +
                comparison_df[DOB_column] + " " +
                comparison_df[name_column]
            )

        return comparison_df
    
#----------------------------------------------------------------------------------------   
    # Function to match names with configurable common part threshold
    def match_names(df_temp, comp_temp, min_common_parts=1):
        df_parts = [part.strip() for part in df_temp.split()]
        comp_parts = [part.strip() for part in comp_temp.split()]

        if df_parts[:2] != comp_parts[:2]:
            return False

        df_name_parts = set(df_parts[2:])
        comp_name_parts = set(comp_parts[2:])

        common_parts = df_name_parts.intersection(comp_name_parts)

        return len(common_parts) >= min_common_parts


    # Map Active status based on matching criteria
    def get_active_status(row):
        matches = comparison_df.loc[comparison_df['Temp_Column'].apply(lambda y: match_names(row['Temp_Column'], y))]
        if not matches.empty:
            if active_column in matches.columns:
                return matches[active_column].values[0]  # Access the active_column safely
            else:
                return None  # Or handle it in some other way if the column doesn't exist
        else:
            return None  # Or handle the case where no match is found

               
    def get_match(df_temp, comparison_df):
        for comp_temp in comparison_df['Temp_Column']:
            if match_names(df_temp, comp_temp):
                return comp_temp
        return None
        
#---------------------------------------------------------------------------------------- 
        

    # Function to match names with configurable common part threshold
    def match_name(df_temp, comp_temp, min_common_parts=3):
        df_parts = [part.strip() for part in df_temp.split()]
        comp_parts = [part.strip() for part in comp_temp.split()]

        if df_parts[0] != comp_parts[0]:
            return False

        df_name_parts = set(df_parts[1:])
        comp_name_parts = set(comp_parts[1:])

        common_parts = df_name_parts.intersection(comp_name_parts)

        return len(common_parts) >= min_common_parts

    # Map Active status based on matching criteria
    def get_active_status_(row):
        matches = comparison_df.loc[comparison_df['Temp_Column'].apply(lambda y: match_name(row['Temp_Column'], y))]
        if not matches.empty:
            if active_column in matches.columns:
                return matches[active_column].values[0]  # Access the active_column safely
            else:
                return None  # Or handle it in some other way if the column doesn't exist
        else:
            return None  # Or handle the case where no match is found
                 
        
    def get_match(df_temp, comparison_df):
        for comp_temp in comparison_df['Temp_Column']:
            if match_name(df_temp, comp_temp):
                return comp_temp
        return None
        
#----------------------------------------------------------------------------------------        

    if not file_paths_C:
        print("Process Type: Addition: No file found for Active Roster \n")
    else:
        comparison_df, comparison_sheets = load_comparison_sheet(file_paths_C[0])
        if comparison_df is None:
            print(f"Workbook does not have GMC, GPA, or GTL sheets. Available sheets: {list(comparison_sheets.keys())} \n")
            comparison_df = get_valid_sheet(comparison_sheets)
            
            if comparison_df is not None:
                # Convert column names to title case
                comparison_df.columns = [col.strip().title() for col in comparison_df.columns]
                for column in comparison_df.select_dtypes(include=['object']).columns:
                    comparison_df[column] = clean_column(comparison_df, column)

                comparison_df.dropna(how='all', inplace=True)
                
                if comparison_df is None:
                    print("Process Type: Addition: No file found for Active Roster \n")
                else:
                    if isinstance(comparison_df, pd.DataFrame) and not comparison_df.empty:
                        while True:
                            name_column, emp_no_column, DOB_column, active_column = validate_columns(comparison_df)
                            if comparison_df[name_column].dropna().empty or comparison_df[emp_no_column].dropna().empty:
                                print(f"{name_column} or {emp_no_column} Columns are empty. \n")
                                print(f"Available Columns: {list(comparison_df.columns)}")
                                user_choice = input("Are you providing another column (yes) or looking for another sheet (no)? ").lower().strip()
                                if user_choice == 'no':
                                    comparison_df = get_valid_sheet(comparison_sheets)
                                    if comparison_df is None:
                                        print("Process Type: Addition: No file found for Active Roster \n")
                                else:
                                    if comparison_df[name_column].dropna().empty:
                                        name_column = input("Please input the column name for 'Name': ").strip().title()
                                    if comparison_df[emp_no_column].dropna().empty:
                                        emp_no_column = input("Please input the column name for 'Emp No': ").strip().title()
                                    if comparison_df[DOB_column].dropna().empty:
                                        DOB_column = input("Please input the column name for 'Date Of Birth': ").strip().title()
                                    if comparison_df[active_column].dropna().empty:
                                        active_column = input("Please input the column name for 'Coverage Status': ").strip().title()
                            else:
                                break
                            
                        

                if comparison_df is not None:
                    # Continue with the processing steps
                    comparison_df = process_comparison_df(comparison_df, name_column, emp_no_column, DOB_column, active_column)
                    
                   # Apply the matching function
                    df['Match Found AR'] = df['Temp_Column'].apply(
                        lambda x: any(match_names(x, y) for y in comparison_df['Temp_Column'])
                    )
                    
                    df['Active AR'] = df.apply(get_active_status, axis=1)

                    df['Found on AR'] = df['Temp_Column'].apply(
                        lambda x: get_match(x, comparison_df) if any(match_names(x, y) for y in comparison_df['Temp_Column']) else None
                    )
                    
                    # Apply the matching function
                    df['Match Found AR1'] = df['Temp_Column'].apply(
                        lambda x: any(match_name(x, y) for y in comparison_df['Temp_Column'])
                    )
                    
                    df['Active AR1'] = df.apply(get_active_status_, axis=1)
                    
                    df['Found on AR1'] = df['Temp_Column'].apply(
                        lambda x: get_match(x, comparison_df) if any(match_name(x, y) for y in comparison_df['Temp_Column']) else None
                    )

                    
                    
                    
                    
     
            
    #------------------------------------------ Color Function Begin -------------------------------------------------------
    
    # Function to apply the light blue background color
    def emp_no_color(row):
        if row['Emp No'] in emp_no_set:
            return ['background-color: lightblue'] * len(row)
        else:
            return [''] * len(row)
        
    # Function to apply conditional formatting
    def age_color(row):
        if pd.isnull(row['Age']):
            return ''
        if row['Relation Type'] in ['Self', 'Employee', 'Wife', 'Spouse', 'Father', 'Mother', 'Parent'] and row['Age'] <= 18:
            return 'background-color: yellow'
        elif row['Relation Type'] in ['Child', 'Daughter', 'Son'] and row['Age'] > 25:
            return 'background-color: yellow'
        return ''

    # Apply the age_color function row-wise and store the results
    styles = df.apply(lambda row: age_color(row), axis=1)

    # Apply the conditional formatting to the DataFrame
    df_styled = df.style.apply(lambda x: ['background-color: yellow' 
                                          if val == 'background-color: yellow' 
                                          else '' 
                                          for val in styles], 
                               subset=['Age']).apply(emp_no_color, axis=1)

    return df_styled, None
