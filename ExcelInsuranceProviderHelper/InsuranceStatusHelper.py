from datetime import datetime
import time
import os
import pandas as pd
from InsuranceStatusHelperEnum import ENROLLMENT_STATUS_ENUM, INSURANCE_FORMAT_ENUM, MATCHING_STATUS_ENUM, PLAN_TYPE_ENUM
from logger import Logger
from PyQt5.QtCore import QThread

# ADP sheet column names
ADP_COMPANY_CODE_COLUMN = "COMPANY CODE"
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

# Employee Life sheet column names
EMPLOYEE_LIFE_FIRST_NAME_COLUMN = "First Name"
EMPLOYEE_LIFE_LAST_NAME_COLUMN = "Last Name"
EMPLOYEE_LIFE_DATE_OF_BIRTH_COLUMN = "Date of Birth"
EMPLOYEE_LIFE_DATE_OF_HIRE_COLUMN = "Date of Hire"
EMPLOYEE_LIFE_TERMINATION_DATE_COLUMN = "Termination Date"
EMPLOYEE_CUSTOMER_NUMBER_COLUMN = "Customer Number"

COMMENT_COLUMN = "Comment"


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
    def __init__(self, adp_file_full_path : str, insurance_file_full_path : str, insurance_provider_type : INSURANCE_FORMAT_ENUM, plan_type : PLAN_TYPE_ENUM, output_folder : str, logger : Logger = None):
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
            self.worker_thread = GenericWorker(self._generate_status_report, self.logger)
            try:
                self.worker_thread.start()
            except Exception as e:
                print(e)
        else:
            self._generate_status_report()

    def _generate_status_report(self):
        self._log_info(f"Job starting...")
        start_time = time.time()
        report_df = self._get_status_report(self.adp_file_path, self.insurance_file_path, self.insurance_provider_type, self.plan_type)
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

        end_time = time.time()
        time_elapsed = end_time - start_time
        self._log_warning(f"Time elapsed: {time_elapsed:.4f} seconds")
        
    def _create_excel_file(self, df : pd.DataFrame, output_file_full_name: str, overwite : bool = True):
        already_exist = os.path.exists(output_file_full_name)
        if already_exist and not overwite:
            self._log_error(f"Failed to create excel file because {output_file_full_name} already exist.")
            return

        self._log_info(f"Creating excel file {output_file_full_name}...")
        df.to_excel(output_file_full_name, index=False)
        self._log_info(f"Excel file {output_file_full_name} created.")

    def _get_status_report(self, adp_file_full_path : str, insurance_file_full_path : str, insurance_format : INSURANCE_FORMAT_ENUM, plan_type : PLAN_TYPE_ENUM):
        if plan_type == PLAN_TYPE_ENUM.DENTAL:
            return None
        elif plan_type == PLAN_TYPE_ENUM.EMPLOYEE_LIFE:
            return self._get_status_report_for_employee_life(adp_file_full_path, insurance_file_full_path, insurance_format)
        elif plan_type == PLAN_TYPE_ENUM.MEDICAL:
            return None
        elif plan_type == PLAN_TYPE_ENUM.VISION:
            return None
        
    def _get_status_report_for_employee_life(self, adp_file_full_path: str, insurance_file_full_path: str, insurance_format: INSURANCE_FORMAT_ENUM) -> pd.DataFrame:
        adp_df = pd.read_excel(adp_file_full_path, header=None) # Assuming only one sheet in the excel file
        header_row_index = adp_df[adp_df.iloc[:, 5] == ADP_HIRE_DATE_COLUMN].index[0] # Find the header row (where column F has "HIRE DATE")
        adp_df = pd.read_excel(adp_file_full_path, header=header_row_index) 

        company_code = insurance_format.get_company_code_string()
        adp_filtered_df = adp_df[adp_df[ADP_COMPANY_CODE_COLUMN].isin(company_code) & (adp_df[ADP_PLAN_TYPE_COLUMN] == PLAN_TYPE_ENUM.EMPLOYEE_LIFE.get_string())]
        adp_filtered_df = adp_filtered_df[[ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN, ADP_HIRE_DATE_COLUMN, ADP_TERMINATION_DATE_COLUMN, ADP_ENROLLMENT_STATUS_COLUMN]] # strip the table and left only the wanted columns
        adp_filtered_df[ADP_DATE_OF_BIRTH_COLUMN] = adp_filtered_df[ADP_DATE_OF_BIRTH_COLUMN].apply(self._get_excel_serial_date)
        adp_filtered_df[ADP_HIRE_DATE_COLUMN] = adp_filtered_df[ADP_HIRE_DATE_COLUMN].apply(self._get_excel_serial_date)
        adp_filtered_df[ADP_TERMINATION_DATE_COLUMN] = adp_filtered_df[ADP_TERMINATION_DATE_COLUMN].apply(self._get_excel_serial_date)

        #merged_df = pd.merge(adp_filtered_df, adp_filtered_df, on=[ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN], how="outer", suffixes=["_Left", "_Right"])

        insurance_df = pd.read_excel(insurance_file_full_path, header=None) # Assuming only one sheet in the excel file
        header_row_index = insurance_df[insurance_df.iloc[:, 2] == BSS_FIRST_NAME_COLUMN].index[0] # Find the header row (where column C has "First Name")
        insurance_df = pd.read_excel(insurance_file_full_path, header=header_row_index)
        insurance_df[ADP_NAME_COLUMN] = insurance_df[EMPLOYEE_LIFE_LAST_NAME_COLUMN] + ", " + insurance_df[EMPLOYEE_LIFE_FIRST_NAME_COLUMN]
        insurance_df = insurance_df.rename(columns={EMPLOYEE_LIFE_DATE_OF_BIRTH_COLUMN : ADP_DATE_OF_BIRTH_COLUMN,
                                                    EMPLOYEE_LIFE_DATE_OF_HIRE_COLUMN : ADP_HIRE_DATE_COLUMN,
                                                    EMPLOYEE_LIFE_TERMINATION_DATE_COLUMN : ADP_TERMINATION_DATE_COLUMN})
        insurance_df = insurance_df[[ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN, ADP_HIRE_DATE_COLUMN, ADP_TERMINATION_DATE_COLUMN, EMPLOYEE_CUSTOMER_NUMBER_COLUMN]]     # strip the table and left only the wanted columns

        #merged_df = pd.merge(insurance_df, insurance_df, on=[ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN], how="outer", suffixes=["_Left", "_Right"])


        merged_df = pd.merge(insurance_df, adp_filtered_df, on=[ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN], how="outer", suffixes=["_insurance", "_ADP"])
        merged_df["Comment"] = ""
        new_insurance_date_of_hire_column = ADP_HIRE_DATE_COLUMN + "_insurance"
        new_insurance_termination_date_column = ADP_TERMINATION_DATE_COLUMN + "_ADP"
        new_adp_date_of_hire_column = ADP_HIRE_DATE_COLUMN + "_ADP"
        new_adp_termination_date_column = ADP_HIRE_DATE_COLUMN + "_ADP"

        new_comment = ""
        for row_index, row in merged_df.iterrows():
            if row["Comment"] != "":
                new_comment = MATCHING_STATUS_ENUM.DUPLICATE_FOUND.get_string()
            elif pd.isna(row[EMPLOYEE_CUSTOMER_NUMBER_COLUMN]):
                new_comment = MATCHING_STATUS_ENUM.EXIST_ONLY_IN_ADP.get_string()
            elif row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.ACTIVE.get_string():
                if row[new_insurance_date_of_hire_column] != row[new_adp_date_of_hire_column]:
                    new_comment = MATCHING_STATUS_ENUM.MISMATCHING_START_DATE.get_string()
                else:
                    new_comment = MATCHING_STATUS_ENUM.GOOD_MATCHING.get_string()
            elif row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.INACTIVE.get_string():
                if row[new_insurance_termination_date_column] != row[new_adp_termination_date_column]:
                    new_comment = MATCHING_STATUS_ENUM.MISMATCHING_END_DATE.get_string()
                else:
                    new_comment = MATCHING_STATUS_ENUM.GOOD_MATCHING.get_string()
            else:
                new_comment = MATCHING_STATUS_ENUM.NEED_TO_BE_IN_ADP.get_string()
            merged_df.at[row_index, "Comment"] = new_comment

        return merged_df
        


    def _get_status_report_for_bss(self, adp_file_full_path : str, bss_file_full_path : str, plan_type: PLAN_TYPE_ENUM) -> pd.DataFrame:
        self._log_info(f"Retriving dataframe from {adp_file_full_path}...")
        adp_df = pd.read_excel(adp_file_full_path, header=None) # Assuming only one sheet in the excel file
        header_row_index = adp_df[adp_df.iloc[:, 5] == ADP_HIRE_DATE_COLUMN].index[0] # Find the header row (where column F has "HIRE DATE")
        adp_df = pd.read_excel(adp_file_full_path, header=header_row_index) 
        self._log_info(f"Done retriving dataframe from {adp_file_full_path}. Row count: {len(adp_df)}.")

        self._log_info(f"Filtering dataframe with plan type {plan_type.get_string()}...")
        adp_df_with_same_plan_type = self._filter_by_columns(adp_df, [ADP_PLAN_TYPE_COLUMN], [plan_type.get_string()]) # keep only row with the given plan_type
        adp_df_with_same_plan_type["Found"] = False
        self._log_info(f"Done filtering. Row count with {plan_type.get_string()}: {len(adp_df_with_same_plan_type)}.")

        self._log_info(f"Retriving dataframe from {bss_file_full_path}...")
        bss_df = pd.read_excel(bss_file_full_path, header=None) # Assuming only one sheet in the excel file
        header_row_index = bss_df[bss_df.iloc[:, 2] == BSS_FIRST_NAME_COLUMN].index[0] # Find the header row (where column C has "First Name")
        bss_df = pd.read_excel(bss_file_full_path, header=header_row_index)
        self._log_info(f"Done retriving dataframe from {bss_file_full_path}. Row count: {len(bss_df)}.")
        bss_df = bss_df[[BSS_FIRST_NAME_COLUMN, BSS_LAST_NAME_COLUMN, BSS_DATE_OF_BIRTH_COLUMN, BSS_DATE_OF_HIRE_COLUMN, BSS_TERMINATION_DATE_COLUMN]]
        bss_df[ADP_PLAN_TYPE_COLUMN] = plan_type.get_string()
        bss_df[COMMENT_COLUMN] = ""
        
        self._log_info("Populating new dataframe with status...")
        seen_rows = set()
        for bss_row_index, bss_row in bss_df.iterrows():
            new_comment_to_add = None

            employee_first_name_bss = bss_row[BSS_FIRST_NAME_COLUMN]
            employee_last_name_bss = bss_row[BSS_LAST_NAME_COLUMN]
            employee_date_of_birth_bss = bss_row[BSS_DATE_OF_BIRTH_COLUMN]
            unique_key = (employee_first_name_bss, employee_last_name_bss, employee_date_of_birth_bss)
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
                        hire_date_bss = bss_row[BSS_DATE_OF_HIRE_COLUMN]
                        if hire_date_adp != hire_date_bss:
                            new_comment_to_add = MATCHING_STATUS_ENUM.MISMATCHING_START_DATE.get_string()
                    elif adp_row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.INACTIVE.get_string():
                        termination_date_adp = self._get_excel_serial_date(adp_row[ADP_TERMINATION_DATE_COLUMN])
                        termination_date_bss = bss_row[BSS_TERMINATION_DATE_COLUMN]
                        if termination_date_adp != termination_date_bss:
                            new_comment_to_add = MATCHING_STATUS_ENUM.MISMATCHING_END_DATE.get_string()
                    adp_df_with_same_plan_type.at[adp_row_index, "Found"] = found_in_adp # assuming an employee with that insurance type would only exist once in ADP
                    break # No need to continue since each employee is unique in ADP

            if not found_in_adp:
                new_comment_to_add = MATCHING_STATUS_ENUM.NEED_TO_BE_IN_ADP.get_string()

            bss_df[COMMENT_COLUMN] = new_comment_to_add

        bss_df.loc[len(bss_df)] = {BSS_FIRST_NAME_COLUMN : "",
                                     BSS_LAST_NAME_COLUMN : "", 
                                BSS_DATE_OF_BIRTH_COLUMN : "", 
                                BSS_DATE_OF_HIRE_COLUMN : "",
                                 BSS_TERMINATION_DATE_COLUMN : "", 
                                 ADP_PLAN_TYPE_COLUMN: "",
                                 COMMENT_COLUMN : ""}  # Append an empty row to the final DataFrame

        for row_index, row in adp_df_with_same_plan_type.iterrows(): # All left over item are not in bss side
            if row["Found"] is False:
                last_name, first_name = self._get_last_and_first_name(row[ADP_NAME_COLUMN])
                date_of_birth = row[ADP_DATE_OF_BIRTH_COLUMN]     
                new_row = {BSS_FIRST_NAME_COLUMN : first_name,
                        BSS_LAST_NAME_COLUMN : last_name, 
                   BSS_DATE_OF_BIRTH_COLUMN : date_of_birth, 
                   BSS_DATE_OF_HIRE_COLUMN : row[ADP_HIRE_DATE_COLUMN],
                    BSS_TERMINATION_DATE_COLUMN : row[ADP_TERMINATION_DATE_COLUMN], 
                    ADP_PLAN_TYPE_COLUMN: plan_type.get_string(),
                    COMMENT_COLUMN : MATCHING_STATUS_ENUM.EXIST_ONLY_IN_ADP.get_string()}
                bss_df = pd.concat([bss_df, pd.DataFrame([new_row])], ignore_index=True) # Append the new row to the final DataFrame

        self._log_info(f"New dataframe populated. Row count: {len(bss_df)}.")
        return bss_df
    
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
        



