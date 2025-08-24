
from datetime import datetime
import os
import shutil
import pandas as pd
from ExcelInsuranceProviderHelper.InsuranceStatusHelperEnum import ENROLLMENT_STATUS_ENUM, INSURANCE_PROVIDER_ENUM, MATCHING_STATUS_ENUM, PLAN_TYPE_ENUM

FILE_LOCATION = "D://sample.xlsx"   # this is used for debugging

# ADP sheet column names
ADP_NAME_COLUMN = "NAME"
ADP_EMPLOYEE_STATUS_COLUMN = "EMPLOYEE STATUS"
ADP_DATE_OF_BIRTH_COLUMN = "DATE OF BIRTH"
ADP_HIRE_DATE_COLUMN = "HIRE DATE"
ADP_TERMINATION_DATE_COLUMN = "TERMINATION DATE"
ADP_PLAN_TYPE_COLUMN = "PLAN TYPE"
ADP_ENROLLMENT_STATUS_COLUMN = "ENROLLMENT STATUS"

# BFS sheet column names
BFS_FIRST_NAME_COLUMN = "First Name"
BFS_LAST_NAME_COLUMN = "Last Name"
BFS_DATE_OF_BIRTH_COLUMN = "Date of Birth"
BFS_DATE_OF_HIRE_COLUMN = "Date of Hire"
BFS_TERMINATION_DATE_COLUMN = "Termination Date"

# Insurance sheet column names. Assuming all insurance providers use the same column names
INSURANCE_NAME_COLUMN = "full name"
INSURANCE_DATE_OF_BIRTH_COLUMN = "Date of Birth"
INSURANCE_DATE_OF_HIRE_COLUMN = "Date of Hire"
INSURANCE_TERMINATION_DATE_COLUMN = "Termination Date"

# Sheet Names
ADP_SHEET_NAME = "Employee Enrollments"
BFS_SHEET_NAME = "bfs"




"""
This is a helper to retrieve insurance status of each employee
"""
class InsuranceStatusHelper:
    def __init__(self, adp_file_full_path : str, insurance_file_full_path : str, insurance_provider_type : INSURANCE_PROVIDER_ENUM, plan_type : PLAN_TYPE_ENUM, output_folder : str):
        self.adp_file_path = adp_file_full_path
        self.insurance_file_path = insurance_file_full_path
        self.plan_type = plan_type
        self.insurance_provider_type = insurance_provider_type
        self.output_folder = output_folder

    def generate_status_report(self):
        report_df = self.get_status_report(self.adp_file_path, self.insurance_file_path, self.insurance_provider_type, self.plan_type, self.output_folder)
        if report_df is None:
            print("Failed to get status report dataframe!")
            return
        
    def create_excel_file(df : pd.DataFrame, output_file_full_name: str, overwite : bool = True):
        already_exist = os.path.exists(output_file_full_name)
        proceed = True
        if already_exist:
            if not overwite:
                proceed = False

        if proceed:
            # TODO: create the output file
            pass

    def get_status_report(self, adp_file_full_path : str, insurance_file_full_path : str, insurance_provider_type : INSURANCE_PROVIDER_ENUM, plan_type : PLAN_TYPE_ENUM, output_folder : str):
        if insurance_provider_type == INSURANCE_PROVIDER_ENUM.CIGNA:
            return None
        elif insurance_provider_type == INSURANCE_PROVIDER_ENUM.BFS:
            return self.get_status_report_for_bfs(adp_file_full_path, insurance_file_full_path, plan_type)
        elif insurance_provider_type == INSURANCE_PROVIDER_ENUM.BSS:
            return None
        else:
            return None

    def get_status_report_for_bfs(self, adp_file_full_path : str, bfs_file_full_path : str, plan_type: PLAN_TYPE_ENUM) -> pd.DataFrame:
        adp_xls = pd.ExcelFile(adp_file_full_path) # read it as excel file first in case there are more than 1 sheet
        adp_df = pd.read_excel(adp_xls, ADP_SHEET_NAME)
        adp_df = self.filter_by_columns(adp_df, [ADP_PLAN_TYPE_COLUMN], [plan_type.get_string()]) # keep only row with the given plan_type

        bfs_xls = pd.ExcelFile(bfs_file_full_path) # read it as excel file first in case there are more than 1 sheet
        bfs_df = pd.read_excel(bfs_xls, BFS_SHEET_NAME)

        final_df = pd.DataFrame(columns = [ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN, ADP_HIRE_DATE_COLUMN, ADP_TERMINATION_DATE_COLUMN, "Comments"])   # Initialize the final DataFrame to store results

        new_comment_key = INSURANCE_PROVIDER_ENUM.BFS.get_string()

        for adp_row_index, adp_row in adp_df.iterrows():
            new_comments = {new_comment_key : ""}
            employee_fullname = adp_row[ADP_NAME_COLUMN] # Format: last, first
            employee_date_of_birth = self.get_excel_serial_date(adp_row[ADP_DATE_OF_BIRTH_COLUMN]) # ADP always have "Month/Day/Year". Need to convert it first
            employee_last_name, employee_first_name = self.get_last_and_first_name(employee_fullname)

            same_employee_in_bfs_df = self.filter_by_columns(bfs_df, [BFS_FIRST_NAME_COLUMN, BFS_LAST_NAME_COLUMN, BFS_DATE_OF_BIRTH_COLUMN], [employee_first_name, employee_last_name, employee_date_of_birth])

            if len(same_employee_in_bfs_df) == 0:
                new_comment_to_add = MATCHING_STATUS_ENUM.NOT_EXIST.get_string()
            else:
                new_comment_to_add = None
                found = False
                if adp_row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.ACTIVE.get_string():
                    for bfs_row_index, bfs_row in same_employee_in_bfs_df.iterrows():
                        adp_hire_date = self.get_excel_serial_date(adp_row[ADP_HIRE_DATE_COLUMN])
                        bfs_hire_date = bfs_row[BFS_DATE_OF_HIRE_COLUMN]
                        if adp_hire_date == bfs_hire_date:
                            if found:
                                new_comment_to_add = MATCHING_STATUS_ENUM.DUPLICATE_FOUND.get_string()
                                break
                            else:
                                new_comment_to_add = MATCHING_STATUS_ENUM.GOOD_MATCHING.get_string()
                                found = True
                    if not found:
                        new_comment_to_add = MATCHING_STATUS_ENUM.MISMATCHING_START_DATE.get_string()
                elif adp_row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.INACTIVE.get_string(): # Use elif instead because cannot assume there are only Active/Inactive 
                    for bfs_row_index, bfs_row in same_employee_in_bfs_df.iterrows():
                            adp_termination_date = self.get_excel_serial_date(adp_row[ADP_TERMINATION_DATE_COLUMN])
                            bfs_termination_date = bfs_row[BFS_TERMINATION_DATE_COLUMN]
                            if adp_termination_date == bfs_termination_date:
                                if found:
                                    new_comment_to_add = MATCHING_STATUS_ENUM.DUPLICATE_FOUND.get_string()
                                    break
                                else:
                                    new_comment_to_add = MATCHING_STATUS_ENUM.GOOD_MATCHING.get_string()
                                    found = True
                    if not found:
                        new_comment_to_add = MATCHING_STATUS_ENUM.MISMATCHING_END_DATE.get_string()

            new_comments[new_comment_key] += new_comment_to_add
                    
            new_row = {ADP_NAME_COLUMN : adp_row[ADP_NAME_COLUMN], 
                   ADP_DATE_OF_BIRTH_COLUMN : adp_row[ADP_DATE_OF_BIRTH_COLUMN], 
                   ADP_HIRE_DATE_COLUMN : adp_row[ADP_HIRE_DATE_COLUMN],
                    ADP_TERMINATION_DATE_COLUMN : adp_row[ADP_TERMINATION_DATE_COLUMN], 
                    ADP_PLAN_TYPE_COLUMN: plan_type.get_string(),
                    "Comments" : new_comments}

            final_df = pd.concat([final_df, pd.DataFrame([new_row])], ignore_index=True) # Append the new row to the final DataFrame
            
        return final_df


    def get_last_and_first_name(self, full_name: str) -> tuple[str, str]:
        """
        full name is in format: "Last, First"
        Return: [Last, First]
        """
        parts = full_name.split(",")
        last_name = parts[0].strip()
        first_name = parts[1].strip()
        return last_name, first_name
    
    def get_excel_serial_date(self, date_string, format: str = "%m/%d/%Y") -> int:
        if not isinstance(date_string, str):
            print(f"{date_string} is not a string.")
            return -1
        try:
            dt = datetime.strptime(date_string, format)
        except TypeError:
            print("Invalid date string")
        excel_start = datetime(1899, 12, 30)  # Excel's epoch start date
        return (dt - excel_start).days
    
    def filter_by_columns(self, df: pd.DataFrame, column_name: list[str], values: list[str]) -> pd.DataFrame:
        """
        Filters a pandas DataFrame by matching specified columns to given values.
        Args:
            df (pd.DataFrame): The DataFrame to filter.
            column_name (list[str]): List of column names to filter by.
            values (list): List of values to match for each corresponding column.
        Returns:
            pd.DataFrame: A filtered DataFrame containing rows where each specified column matches the corresponding value.
        Raises:
            ValueError: If the length of column_name and values lists do not match.
        Example:
            >>> filter_by_columns(df, ['status', 'type'], ['active', 'premium'])
            Returns all rows where df['status'] == 'active' and df['type'] == 'premium'.
        """
        if len(column_name) != len(values):
            raise ValueError("columns and values must have the same length")

        # Start with a mask of all True
        mask = pd.Series([True] * len(df))

        for col, val in zip(column_name, values):
            mask &= (df[col] == val)  # Combine conditions

        return df[mask]
        



