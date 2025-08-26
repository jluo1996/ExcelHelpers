from datetime import datetime
import os
import pandas as pd
from AdpEmployee import AdpEmployee
from InsuranceStatusHelperEnum import ENROLLMENT_STATUS_ENUM, INSURANCE_PROVIDER_ENUM, MATCHING_STATUS_ENUM, PLAN_TYPE_ENUM
from logger import Logger
from PyQt5.QtCore import QThread

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

# BSS sheet column names
BSS_FIRST_NAME_COLUMN = "First Name"
BSS_LAST_NAME_COLUMN = "Last Name"
BSS_DATE_OF_BIRTH_COLUMN = "Date of Birth"
BSS_DATE_OF_HIRE_COLUMN = "Date of Hire"
BSS_TERMINATION_DATE_COLUMN = "Termination Date"

# Insurance sheet column names. Assuming all insurance providers use the same column names
INSURANCE_NAME_COLUMN = "full name"
INSURANCE_DATE_OF_BIRTH_COLUMN = "Date of Birth"
INSURANCE_DATE_OF_HIRE_COLUMN = "Date of Hire"
INSURANCE_TERMINATION_DATE_COLUMN = "Termination Date"

# Sheet Names
ADP_SHEET_NAME = "Employee Enrollments"
BFS_SHEET_NAME = "bfs"
BSS_SHEET_NAME = "bss"

class GenericWorker(QThread):
    def __init__(self, func, logger : Logger = None):
        super().__init__()

        self.func = func
        self.logger = logger

    def run(self):
        try:
            self.func() 
        except Exception as e:
            if self.logger is not None:
                self.logger.log_error(f"Exception is caught: {e}")



"""
This is a helper to retrieve insurance status of each employee
"""
class InsuranceStatusHelper:
    def __init__(self, adp_file_full_path : str, insurance_file_full_path : str, insurance_provider_type : INSURANCE_PROVIDER_ENUM, plan_type : PLAN_TYPE_ENUM, output_folder : str, logger : Logger = None):
        self.adp_file_path = adp_file_full_path
        self.insurance_file_path = insurance_file_full_path
        self.plan_type = plan_type
        self.insurance_provider_type = insurance_provider_type
        self.output_folder = output_folder
        self.logger = logger

    def _log_info(self, msg):
        if self.logger:
            self.logger.log_info(msg)
        else:
            print(msg)

    def _log_error(self, msg):
        if self.logger:
            self.logger.log_error(msg)
        else:
            print(msg)

    def _log_warning(self, msg):
        if self.logger:
            self.logger.log_warning(msg)
        else:
            print(msg)


    def generate_status_report(self, run_as_thread = False):
        if run_as_thread:
            self.worker_thread = GenericWorker(self._generate_status_report)
            self.worker_thread.start()
        else:
            self._generate_status_report()

    def _generate_status_report(self, run_as_thread = False):
        report_df = self._get_status_report(self.adp_file_path, self.insurance_file_path, self.insurance_provider_type, self.plan_type, self.output_folder)
        if report_df is None:
            self._log_error(f"{self.insurance_provider_type.get_string()} is not supported.")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file_name = f"StatusReport_{self.insurance_provider_type.get_string()}_{self.plan_type.get_string()}_{timestamp}.xlsx"
        output_file_full_path = os.path.join(self.output_folder, output_file_name)
        self._create_excel_file(report_df, output_file_full_path)

        # Create .csv file as well for debug
        if __debug__:
            csv_file_path = output_file_full_path.replace(".xlsx", '.csv')
            report_df.to_csv(csv_file_path, index=False)

        # --------- end of debug code
        
    def _create_excel_file(self, df : pd.DataFrame, output_file_full_name: str, overwite : bool = True):
        already_exist = os.path.exists(output_file_full_name)
        if already_exist and not overwite:
            self._log_error(f"Failed to create excel file because {output_file_full_name} already exist.")
            return

        self._log_info(f"Creating excel file {output_file_full_name}...")
        df.to_excel(output_file_full_name, index=False)
        self._log_info(f"Excel file {output_file_full_name} created.")

    def _get_status_report(self, adp_file_full_path : str, insurance_file_full_path : str, insurance_provider_type : INSURANCE_PROVIDER_ENUM, plan_type : PLAN_TYPE_ENUM, output_folder : str):
        if insurance_provider_type == INSURANCE_PROVIDER_ENUM.CIGNA:
            return None
        elif insurance_provider_type == INSURANCE_PROVIDER_ENUM.BFS:
            return self._get_status_report_for_bfs(adp_file_full_path, insurance_file_full_path, plan_type)
        elif insurance_provider_type == INSURANCE_PROVIDER_ENUM.BSS:
            return self._get_status_report_for_bfs(adp_file_full_path, insurance_file_full_path, plan_type) # use the BFS method since format and logic is the same
        else:
            return None
        
    def _get_status_report_for_bss(self, adp_file_full_path : str, bss_file_full_path : str, plan_type: PLAN_TYPE_ENUM) -> pd.DataFrame:
        adp_df = pd.read_excel(adp_file_full_path, header=None) # Assuming only one sheet in the excel file
        header_row_index = adp_df[adp_df.iloc[:, 5] == ADP_HIRE_DATE_COLUMN].index[0] # Find the header row (where column F has "HIRE DATE")
        adp_df = pd.read_excel(adp_file_full_path, header=header_row_index) 

        # Convert dataframe to list of ADP employees
        adp_employee_list = []
        for row_index, row in adp_df.iterrows():
            adp_employee = AdpEmployee(full_name=row[ADP_NAME_COLUMN],
                                       date_of_birth=self._get_excel_serial_date(row[ADP_DATE_OF_BIRTH_COLUMN]),
                                       date_of_hire=self._get_excel_serial_date(row[ADP_HIRE_DATE_COLUMN]),
                                       termination_date=self._get_excel_serial_date(row[ADP_TERMINATION_DATE_COLUMN]),
                                       )

    def _get_status_report_for_bfs(self, adp_file_full_path : str, bfs_file_full_path : str, plan_type: PLAN_TYPE_ENUM) -> pd.DataFrame:
        self._log_info(f"Retriving dataframe from {adp_file_full_path}...")
        adp_df = pd.read_excel(adp_file_full_path, header=None) # Assuming only one sheet in the excel file
        header_row_index = adp_df[adp_df.iloc[:, 5] == ADP_HIRE_DATE_COLUMN].index[0] # Find the header row (where column F has "HIRE DATE")
        adp_df = pd.read_excel(adp_file_full_path, header=header_row_index) 
        self._log_info(f"Finish retriving dataframe from {adp_file_full_path}. Row count: {len(adp_df)}.")

        self._log_info(f"Filtering rows with only {plan_type.get_string()}...")
        adp_df_with_same_plan_type = self._filter_by_columns(adp_df, [ADP_PLAN_TYPE_COLUMN], [plan_type.get_string()]) # keep only row with the given plan_type
        self._log_info(f"Done filtering. Row count: {len(adp_df_with_same_plan_type)}.")

        self._log_info(f"Retriving dataframe from {bfs_file_full_path}...")
        bfs_df = pd.read_excel(bfs_file_full_path, header=None) # Assuming only one sheet in the excel file
        header_row_index = bfs_df[bfs_df.iloc[:, 2] == BFS_FIRST_NAME_COLUMN].index[0] # Find the header row (where column C has "First Name")
        bfs_df = pd.read_excel(bfs_file_full_path, header=header_row_index)
        self._log_info(f"Finish retriving dataframe from {bfs_file_full_path}. Row count: {len(bfs_df)}.")

        final_df = pd.DataFrame(columns = [BFS_FIRST_NAME_COLUMN, BFS_LAST_NAME_COLUMN, BFS_DATE_OF_BIRTH_COLUMN, BFS_DATE_OF_HIRE_COLUMN, BFS_TERMINATION_DATE_COLUMN, ADP_PLAN_TYPE_COLUMN ,"Comments"])   # Initialize the final DataFrame to store results
        
        self._log_info("Populating new dataframe with status...")
        seen_rows = set()
        for bfs_row_index, bfs_row in bfs_df.iterrows():
            new_comment_to_add = None

            employee_first_name_bfs = bfs_row[BFS_FIRST_NAME_COLUMN]
            employee_last_name_bfs = bfs_row[BFS_LAST_NAME_COLUMN]
            employee_date_of_birth_bfs = bfs_row[BFS_DATE_OF_BIRTH_COLUMN]
            unique_key = (employee_first_name_bfs, employee_last_name_bfs, employee_date_of_birth_bfs)
            if unique_key in seen_rows:
                new_comment_to_add = MATCHING_STATUS_ENUM.DUPLICATE_FOUND.get_string()
            else:
                seen_rows.add(unique_key)

            found_in_adp = False
            
            for adp_row_index, adp_row in adp_df_with_same_plan_type.iterrows():
                employee_full_name_adp = adp_row[ADP_NAME_COLUMN]
                employee_last_name_adp, employee_first_name_adp  = self._get_last_and_first_name(employee_full_name_adp)
                employee_date_of_birth_adp = self._get_excel_serial_date(adp_row[ADP_DATE_OF_BIRTH_COLUMN]) # ADP always have "Month/Day/Year". Need to convert it first
                
                if unique_key == (employee_first_name_adp, employee_last_name_adp, employee_date_of_birth_adp): # Found it at adp
                    found_in_adp = True
                    new_comment_to_add = MATCHING_STATUS_ENUM.GOOD_MATCHING.get_string()

                    if adp_row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.ACTIVE.get_string():
                        hire_date_adp = self._get_excel_serial_date(adp_row[ADP_HIRE_DATE_COLUMN])
                        hire_date_bfs = bfs_row[BFS_DATE_OF_HIRE_COLUMN]
                        if hire_date_adp != hire_date_bfs:
                            new_comment_to_add = MATCHING_STATUS_ENUM.MISMATCHING_START_DATE.get_string()
                    elif adp_row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.INACTIVE.get_string():
                        termination_date_adp = self._get_excel_serial_date(adp_row[ADP_TERMINATION_DATE_COLUMN])
                        termination_date_bfs = bfs_row[BFS_TERMINATION_DATE_COLUMN]
                        if termination_date_adp != termination_date_bfs:
                            new_comment_to_add = MATCHING_STATUS_ENUM.MISMATCHING_END_DATE.get_string()
                    break # No need to continue since each employee is unique in ADP

            if not found_in_adp:
                new_comment_to_add = MATCHING_STATUS_ENUM.NEED_TO_BE_IN_ADP.get_string()

            new_row = {BFS_FIRST_NAME_COLUMN : bfs_row[BFS_FIRST_NAME_COLUMN],
                        BFS_LAST_NAME_COLUMN : bfs_row[BFS_LAST_NAME_COLUMN], 
                   BFS_DATE_OF_BIRTH_COLUMN : bfs_row[BFS_DATE_OF_BIRTH_COLUMN], 
                   BFS_DATE_OF_HIRE_COLUMN : bfs_row[BFS_DATE_OF_HIRE_COLUMN],
                    BFS_TERMINATION_DATE_COLUMN : bfs_row[BFS_TERMINATION_DATE_COLUMN], 
                    ADP_PLAN_TYPE_COLUMN: plan_type.get_string(),
                    "Comments" : new_comment_to_add}
            final_df = pd.concat([final_df, pd.DataFrame([new_row])], ignore_index=True) # Append the new row to the final DataFrame

        new_row = {BFS_FIRST_NAME_COLUMN : "",
                        BFS_LAST_NAME_COLUMN : "", 
                   BFS_DATE_OF_BIRTH_COLUMN : "", 
                   BFS_DATE_OF_HIRE_COLUMN : "",
                    BFS_TERMINATION_DATE_COLUMN : "", 
                    ADP_PLAN_TYPE_COLUMN: "",
                    "Comments" : ""}
        final_df = pd.concat([final_df, pd.DataFrame([new_row])], ignore_index=True) # Append an empty row to the final DataFrame

        for row_index, row in adp_df_with_same_plan_type.iterrows():
             last_name, first_name = self._get_last_and_first_name(row[ADP_NAME_COLUMN])
             date_of_birth = row[ADP_DATE_OF_BIRTH_COLUMN]
             unique_key = (first_name, last_name, date_of_birth)

             if unique_key not in seen_rows:
                # Doesn't exist in the bfs side
                new_row = {BFS_FIRST_NAME_COLUMN : first_name,
                        BFS_LAST_NAME_COLUMN : last_name, 
                   BFS_DATE_OF_BIRTH_COLUMN : date_of_birth, 
                   BFS_DATE_OF_HIRE_COLUMN : row[ADP_HIRE_DATE_COLUMN],
                    BFS_TERMINATION_DATE_COLUMN : row[ADP_TERMINATION_DATE_COLUMN], 
                    ADP_PLAN_TYPE_COLUMN: plan_type.get_string(),
                    "Comments" : MATCHING_STATUS_ENUM.EXIST_ONLY_IN_ADP.get_string()}
                final_df = pd.concat([final_df, pd.DataFrame([new_row])], ignore_index=True) # Append the new row to the final DataFrame

        self._log_info(f"New dataframe populated. Row count: {len(final_df)}.")
        return final_df
    
    def _get_last_and_first_name(self, full_name: str) -> tuple[str, str]:
        """
        full name is in format: "Last, First"
        Return: [Last, First]
        """
        parts = full_name.split(",")
        last_name = parts[0].strip()
        first_name = parts[1].strip()
        return last_name, first_name
    
    def _get_excel_serial_date(self, date_string, format: str = "%m/%d/%Y") -> int:
        if not isinstance(date_string, str):
            #print(f"{date_string} is not a string.")
            return -1
        try:
            dt = datetime.strptime(date_string, format)
        except TypeError:
            print("Invalid date string")
        excel_start = datetime(1899, 12, 30)  # Excel's epoch start date
        return (dt - excel_start).days
    
    def _filter_by_columns(self, df: pd.DataFrame, column_name: list[str], values: list[str]) -> pd.DataFrame:
        final_df = pd.DataFrame(columns=df.columns)
        for row_index, row in df.iterrows():
            matching = True
            for i in range(len(column_name)):
                name = row[column_name[i]]
                value = values[i]
                if name != value:
                    matching = False
            
            if matching:
                final_df = pd.concat([final_df, pd.DataFrame([row])], ignore_index=True)

        return final_df
    

# helper = InsuranceStatusHelper(None, None, None, None, None, None)
# # Example DataFrame
# df = pd.DataFrame({
#     "Name": ["Alice", "Bob"],
#     "Age": [25, 30],
#     "City": ["NY", "LA"]
# })
# output = helper._filter_by_columns(df, ["Name", "Age"], ["Alice", 25])
# print(output)
        



