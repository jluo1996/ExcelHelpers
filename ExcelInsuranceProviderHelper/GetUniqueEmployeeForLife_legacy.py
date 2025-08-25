from datetime import datetime
from enum import Enum
import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog

FILE_LOCATION = "D://sample.xlsx"   # this is used for debugging
OUTPUT_FILE_LOCATION = "D://output.xlsx"    # this is used for debugging

# Sheet names
ADP_SHEET_NAME = "Employee Enrollments"
BFS_SHEET_NAME = "bfs"
BSS_SHEET_NAME = "bss"

# ADP sheet column names
ADP_NAME_COLUMN = "NAME"
ADP_EMPLOYEE_STATUS_COLUMN = "EMPLOYEE STATUS"
ADP_DATE_OF_BIRTH_COLUMN = "DATE OF BIRTH"
ADP_HIRE_DATE_COLUMN = "HIRE DATE"
ADP_TERMINATION_DATE_COLUMN = "TERMINATION DATE"
ADP_PLAN_TYPE_COLUMN = "PLAN TYPE"
ADP_ENROLLMENT_STATUS_COLUMN = "ENROLLMENT STATUS"

# BFS sheet column names
BFS_NAME_COLUMN = "full name"
BFS_DATE_OF_BIRTH_COLUMN = "Date of Birth"
BFS_DATE_OF_HIRE_COLUMN = "Date of Hire"
BFS_TERMINATION_DATE_COLUMN = "Termination Date"

# BSS sheet column names
BSS_NAME_COLUMN = "full"
BSS_DATE_OF_BIRTH_COLUMN = "Date of Birth"
BSS_DATE_OF_HIRE_COLUMN = "Date of Hire"
BSS_TERMINATION_DATE_COLUMN = "Termination Date"

class PLAN_TYPE_ENUM(Enum):
    DENTAL = "Dental"
    EMPLOYEE_LIFE = "Employee Life"
    MEDICAL = "Medical"
    VISION = "Vision"

class ENROLLMENT_STATUS_ENUM(Enum):
    ACTIVE = "Active"
    INACTIVE = "Inactive"

class MATCHING_STATUS_ENUM(Enum):
    GOOD_MATCHING = "Good Matching"
    DUPLICATE_FOUND = "Duplicate Found"
    MISMATCHING_START_DATE = "Mismatching Start Date"
    MISMATCHING_END_DATE = "Mismatching End Date"

class COMPANY_CODE_ENUM(Enum):
    BFS = "E9Y"
    BSS_1 = "E30"
    BSS_2 = "E3V"

class SHEET_NAME_ENUM(Enum):
    ADP = "Employee Enrollments"
    BFS = "bfs"
    BSS = "bss"
    NONE = "None"   # not found in either bfs or bss

'''
Parameters:

date_string: The date as a string.
format: The format of the date string (default is U.S. style month/day/year).
Returns:

An integer representing the Excel serial date.
'''
def get_excel_serial_date(date_string: str, format: str = "%m/%d/%Y") -> int:
    dt = datetime.strptime(date_string, format)
    excel_start = datetime(1899, 12, 30)  # Excel's epoch start date
    return (dt - excel_start).days

# date_of_birth_serial is in Excel serial date format
def get_matching_employee_row_indices(df: pd.DataFrame, name_column:str, date_of_birth_column:str, employee_name: str, date_of_birth_serial: int) -> list[int]:
    return get_matching_row_indices(df, [name_column, date_of_birth_column], [employee_name, date_of_birth_serial])

# Generic function to get matching row indices based on multiple column names and their corresponding values
def get_matching_row_indices(df: pd.DataFrame, column_names: list[str], values: list[str]) -> list[int]:
    if len(column_names) != len(values):
        raise ValueError("Length of column_names and values must be the same.")
    
    condition = pd.Series([True] * len(df))
    for col, val in zip(column_names, values):
        condition &= (df[col] == val)
    
    return df[condition].index.tolist()

def get_excel_file_path():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path


file_path = get_excel_file_path()
if not file_path:
    quit() # Exit if no file is selected
#xls = pd.ExcelFile(FILE_LOCATION)
xls = pd.ExcelFile(file_path)

# Read Excel file
ADP_df = pd.read_excel(xls, ADP_SHEET_NAME)
bfs_df = pd.read_excel(xls, BFS_SHEET_NAME)
bss_df = pd.read_excel(xls, BSS_SHEET_NAME)

final_df = pd.DataFrame(columns = [ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN, ADP_HIRE_DATE_COLUMN, ADP_TERMINATION_DATE_COLUMN, "Comments"])   # Initialize the final DataFrame to store results

employee_with_life_indices = ADP_df[(ADP_df[ADP_PLAN_TYPE_COLUMN] == PLAN_TYPE_ENUM.EMPLOYEE_LIFE.value)].index.tolist()

for index, row in ADP_df.iterrows():
    if row[ADP_PLAN_TYPE_COLUMN] == PLAN_TYPE_ENUM.EMPLOYEE_LIFE.value: # if this employee has life insurance
        new_comment = {} # this dictionary will contain comments for this employee
    
        # find this employee in bfs or bss
        employee_name = row[ADP_NAME_COLUMN]
        employee_date_of_birth_in_excel_serial_date = get_excel_serial_date(row[ADP_DATE_OF_BIRTH_COLUMN]) # This line is only needed if the date in ADP is in string format and needs to be converted to Excel serial date format
        matching_index_list_in_bfs = get_matching_employee_row_indices(bfs_df, BFS_NAME_COLUMN, BFS_DATE_OF_BIRTH_COLUMN, employee_name, employee_date_of_birth_in_excel_serial_date)
        matching_index_list_in_bss = get_matching_employee_row_indices(bss_df, BSS_NAME_COLUMN, BSS_DATE_OF_BIRTH_COLUMN, employee_name, employee_date_of_birth_in_excel_serial_date)
        # -------------------- end of finding this employee in bfs or bss --------------------

        # Identify which sheet was found
        exist_in_bfs = False
        exist_in_bss = False
        if len(matching_index_list_in_bfs) > 0:
            new_comment[SHEET_NAME_ENUM.BFS.value] = ""
            exist_in_bfs = True
        if len(matching_index_list_in_bss) > 0:
            new_comment[SHEET_NAME_ENUM.BSS.value] = ""
            exist_in_bss = True
        if not exist_in_bfs and not exist_in_bss:
            new_comment[SHEET_NAME_ENUM.NONE.value] = "" # not found in either bfs or bss
        else:
            if row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.ACTIVE.value:    # if this employee's life insurance is active in ADP
                # check whether the hire date match in bfs or bss
                if exist_in_bfs:
                    found = False
                    for index in matching_index_list_in_bfs:
                        # check if hire date matches
                        hire_date_ADP = row[ADP_HIRE_DATE_COLUMN]
                        if bfs_df.at[index, BFS_DATE_OF_HIRE_COLUMN] == hire_date_ADP:
                            if found:
                                new_comment[SHEET_NAME_ENUM.BFS.value] += MATCHING_STATUS_ENUM.DUPLICATE_FOUND.value
                                break # no need to check further since duplication is confirmed
                            else:
                                new_comment[SHEET_NAME_ENUM.BFS.value] += MATCHING_STATUS_ENUM.GOOD_MATCHING.value
                                found = True
                    if not found:
                        new_comment[SHEET_NAME_ENUM.BFS.value] += MATCHING_STATUS_ENUM.MISMATCHING_START_DATE.value

                if exist_in_bss:
                    found = False
                    for index in matching_index_list_in_bss:
                        # check if hire date matches
                        hire_date_ADP = row[ADP_HIRE_DATE_COLUMN]
                        if bss_df.at[index, BSS_DATE_OF_HIRE_COLUMN] == hire_date_ADP:
                            if found:
                                new_comment[SHEET_NAME_ENUM.BSS.value] += MATCHING_STATUS_ENUM.DUPLICATE_FOUND.value
                                break
                            else:
                                new_comment[SHEET_NAME_ENUM.BSS.value] += MATCHING_STATUS_ENUM.GOOD_MATCHING.value
                                found = True
                    if not found:
                        new_comment[SHEET_NAME_ENUM.BSS.value] += MATCHING_STATUS_ENUM.MISMATCHING_START_DATE.value
            elif row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.INACTIVE.value:    # Use 'else' if assume Active/Inactive are the only status
                # check whether termination date match in bfs or bss
                if exist_in_bfs:
                    found = False
                    for index in matching_index_list_in_bfs:
                        # check if termination date matches
                        termination_date_ADP = row[ADP_TERMINATION_DATE_COLUMN]
                        if bfs_df.at[index, BFS_TERMINATION_DATE_COLUMN] == termination_date_ADP:
                            if found:
                                new_comment[SHEET_NAME_ENUM.BFS.value] += MATCHING_STATUS_ENUM.DUPLICATE_FOUND.value
                                break # no need to check further since duplication is confirmed
                            else:
                                new_comment[SHEET_NAME_ENUM.BFS.value] += MATCHING_STATUS_ENUM.GOOD_MATCHING.value
                                found = True
                    if not found:
                        new_comment[SHEET_NAME_ENUM.BFS.value] += MATCHING_STATUS_ENUM.MISMATCHING_END_DATE.value

                if exist_in_bss:
                    found = False
                    for index in matching_index_list_in_bss:
                        # check if termination date matches
                        termination_date_ADP = row[ADP_TERMINATION_DATE_COLUMN]
                        if bss_df.at[index, BSS_TERMINATION_DATE_COLUMN] == termination_date_ADP:
                            if found:
                                new_comment[SHEET_NAME_ENUM.BSS.value] += MATCHING_STATUS_ENUM.DUPLICATE_FOUND.value
                                break
                            else:
                                new_comment[SHEET_NAME_ENUM.BSS.value] += MATCHING_STATUS_ENUM.GOOD_MATCHING.value
                                found = True
                    if not found:
                        new_comment[SHEET_NAME_ENUM.BSS.value] += MATCHING_STATUS_ENUM.MISMATCHING_END_DATE.value\
                    
        new_row = {ADP_NAME_COLUMN : row[ADP_NAME_COLUMN], 
                   ADP_DATE_OF_BIRTH_COLUMN : row[ADP_DATE_OF_BIRTH_COLUMN], 
                   ADP_HIRE_DATE_COLUMN : row[ADP_HIRE_DATE_COLUMN],
                    ADP_TERMINATION_DATE_COLUMN : row[ADP_TERMINATION_DATE_COLUMN], 
                    "Comments" : new_comment}
        final_df = pd.concat([final_df, pd.DataFrame([new_row])], ignore_index=True) # append the row to final_df
        row_count = len(final_df)
        print(row_count)

output_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output.xlsx") # output folder path is the same as this script file path
# final_df.to_excel(OUTPUT_FILE_LOCATION, index=False)
# final_df.to_csv(OUTPUT_FILE_LOCATION.replace('.xlsx', '.csv'), index=False)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
if os.path.exists(output_file_path):
    back_up_file_path = output_file_path.replace('.xlsx', f'_backup_{timestamp}.xlsx')
    shutil.copy(output_file_path, back_up_file_path)
final_df.to_excel(output_file_path, index=False)

# Also generate CSV file for debugging purpose
csv_file_path = output_file_path.replace('.xlsx', '.csv')
if os.path.exists(csv_file_path):
    back_up_csv_file_path = csv_file_path.replace('.csv', f'_backup_{timestamp}.csv')
    shutil.copy(csv_file_path, back_up_csv_file_path)
final_df.to_csv(output_file_path.replace('.xlsx', '.csv'), index=False)
# --------------- end of debug code ---------------

        


