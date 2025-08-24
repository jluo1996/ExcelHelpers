from datetime import datetime
import os
import shutil
import pandas as pd

from InsuranceStatusHelperEnum import ENROLLMENT_STATUS_ENUM, MATCHING_STATUS_ENUM, PLAN_TYPE_ENUM

FILE_LOCATION = "D://sample.xlsx"   # this is used for debugging

# ADP sheet column names
ADP_NAME_COLUMN = "NAME"
ADP_EMPLOYEE_STATUS_COLUMN = "EMPLOYEE STATUS"
ADP_DATE_OF_BIRTH_COLUMN = "DATE OF BIRTH"
ADP_HIRE_DATE_COLUMN = "HIRE DATE"
ADP_TERMINATION_DATE_COLUMN = "TERMINATION DATE"
ADP_PLAN_TYPE_COLUMN = "PLAN TYPE"
ADP_ENROLLMENT_STATUS_COLUMN = "ENROLLMENT STATUS"

# Insurance sheet column names. Assuming all insurance providers use the same column names
INSURANCE_NAME_COLUMN = "full name"
INSURANCE_DATE_OF_BIRTH_COLUMN = "Date of Birth"
INSURANCE_DATE_OF_HIRE_COLUMN = "Date of Hire"
INSURANCE_TERMINATION_DATE_COLUMN = "Termination Date"




class InsuranceStatusHelper:
    def __init__(self):
        self.ADP_df = None   # a single dataframe from ADP export 
        self.insurance_df_dict = {}  # a list of dataframe from insurance providers

    def run(self):
        # TODO: ask user to select the file
        self.ADP_df, self.insurance_dfs_dict = self.get_dataframes(FILE_LOCATION)

        # TODO: ask user to select the plan type column name
        plan_type = PLAN_TYPE_ENUM.EMPLOYEE_LIFE
        employee_with_selected_plan_type_df = self.get_employees_by_plan_type(self.ADP_df, ADP_PLAN_TYPE_COLUMN, plan_type)

        final_df = pd.DataFrame(columns = [ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN, ADP_HIRE_DATE_COLUMN, ADP_TERMINATION_DATE_COLUMN, "Comments"])   # Initialize the final DataFrame to store results

        for row_index, row in employee_with_selected_plan_type_df.iterrows():
            new_comments = {}
            employee_name = row[ADP_NAME_COLUMN]
            employee_date_of_birth = self.get_excel_serial_date(row[ADP_DATE_OF_BIRTH_COLUMN])

            for sheet_name, insurance_df in self.insurance_dfs_dict.items():
                matching_employee_df = self.filter_by_columns(insurance_df, [INSURANCE_NAME_COLUMN, INSURANCE_DATE_OF_BIRTH_COLUMN], [employee_name, employee_date_of_birth])

                if len(matching_employee_df) == 0:
                    new_comments[sheet_name] = MATCHING_STATUS_ENUM.NOT_EXIST.value
                    continue

                
                new_comments[sheet_name] = ""
                new_comment_to_add = None
                have_seen = False
                if row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.ACTIVE.value:
                    for match_row_index, match_row in matching_employee_df.iterrows():
                        hire_date_ADP = row[ADP_HIRE_DATE_COLUMN]
                        hire_date_insurance = match_row[INSURANCE_DATE_OF_HIRE_COLUMN]
                        if hire_date_ADP == hire_date_insurance:
                            if have_seen:
                                new_comment_to_add = MATCHING_STATUS_ENUM.DUPLICATE_FOUND.value
                                break
                            else:
                                new_comment_to_add = MATCHING_STATUS_ENUM.GOOD_MATCHING.value
                                have_seen = True
                    if not have_seen:
                        new_comment_to_add = MATCHING_STATUS_ENUM.MISMATCHING_START_DATE.value
                elif row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.INACTIVE.value:
                    for match_row_index, match_row in matching_employee_df.iterrows():
                        termination_date_ADP = row[ADP_TERMINATION_DATE_COLUMN]
                        termination_date_insurance = match_row[INSURANCE_TERMINATION_DATE_COLUMN]
                        if termination_date_ADP == termination_date_insurance:
                            if have_seen:
                                new_comment_to_add = MATCHING_STATUS_ENUM.DUPLICATE_FOUND.value
                                break
                            else:
                                new_comment_to_add = MATCHING_STATUS_ENUM.GOOD_MATCHING.value
                                have_seen = True

                    if not have_seen:
                        new_comment_to_add = MATCHING_STATUS_ENUM.MISMATCHING_END_DATE.value

                new_comments[sheet_name] += new_comment_to_add 

            new_row = {ADP_NAME_COLUMN : row[ADP_NAME_COLUMN], 
                   ADP_DATE_OF_BIRTH_COLUMN : row[ADP_DATE_OF_BIRTH_COLUMN], 
                   ADP_HIRE_DATE_COLUMN : row[ADP_HIRE_DATE_COLUMN],
                    ADP_TERMINATION_DATE_COLUMN : row[ADP_TERMINATION_DATE_COLUMN], 
                    "Comments" : new_comments}

            final_df = pd.concat([final_df, pd.DataFrame([new_row])], ignore_index=True) # Append the new row to the final DataFrame
        
        print(f"Row count of final df: {len(final_df)}")

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


    # TODO: support multiple ADP sheets
    def get_dataframes(self, file_full_path, sheet_index_for_ADP=0):
        '''
        This will populate all the dataframes from the given excel file path.
        The sheet_index_for_ADP is used to identify which sheet is the ADP export sheet.
        '''
        ADP_df = None
        insurance_df = {}
        
        xls = pd.ExcelFile(file_full_path)
        sheet_names = xls.sheet_names
        for idx, sheet_name in enumerate(sheet_names):
            df = pd.read_excel(xls, sheet_name=sheet_name)
            if idx == sheet_index_for_ADP:
                ADP_df = df
            else:
                insurance_df[sheet_name] = df

        return ADP_df, insurance_df
    
    def get_excel_serial_date(self, date_string: str, format: str = "%m/%d/%Y") -> int:
        dt = datetime.strptime(date_string, format)
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
        
    def get_employees_by_plan_type(self, df: pd.DataFrame, plan_type_column_name : str, plan_type: PLAN_TYPE_ENUM) -> pd.DataFrame:
        return self.filter_by_columns(df, [plan_type_column_name], [plan_type.value])



if __name__ == "__main__":
    insurance_status_helper = InsuranceStatusHelper()
    insurance_status_helper.run()
    quit()
