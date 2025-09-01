from datetime import datetime
from pathlib import Path
import time
import os
import pandas as pd
import re
from InsuranceStatusHelperEnum import CIGNA_ID_RELATIONSHIP_ENUM, EMPLOYEE_STATUS_ENUM, ENROLLMENT_STATUS_ENUM, INSURANCE_FORMAT_ENUM, MATCHING_STATUS_ENUM, PLAN_TYPE_ENUM
from logger import Logger
from PyQt5.QtCore import QThread, QUrl

# ADP sheet column names
ADP_COMPANY_CODE_COLUMN = "COMPANY CODE"
ADP_NAME_COLUMN = "NAME"
ADP_TAX_ID_COLUMN = "TAX ID"
ADP_EMPLOYEE_STATUS_COLUMN = "EMPLOYEE STATUS"
ADP_DATE_OF_BIRTH_COLUMN = "DATE OF BIRTH"
ADP_HIRE_DATE_COLUMN = "HIRE DATE"
ADP_TERMINATION_DATE_COLUMN = "TERMINATION DATE"
ADP_PLAN_TYPE_COLUMN = "PLAN TYPE"
ADP_ENROLLMENT_STATUS_COLUMN = "ENROLLMENT STATUS"
ADP_EMPLOYEE_STATUS_COLUMN = "EMPLOYEE STATUS"
ADP_COVERAGE_LEVEL_VALUE_COLUMN = "COVERAGE LEVEL VALUE"
ADP_PROVIDER_COLUMN = "PROVIDER"

# Employee Life sheet column names
EMPLOYEE_LIFE_FIRST_NAME_COLUMN = "First Name"
EMPLOYEE_LIFE_LAST_NAME_COLUMN = "Last Name"
EMPLOYEE_LIFE_DATE_OF_BIRTH_COLUMN = "Date of Birth"
EMPLOYEE_LIFE_DATE_OF_HIRE_COLUMN = "Date of Hire"
EMPLOYEE_LIFE_TERMINATION_DATE_COLUMN = "Termination Date"
EMPLOYEE_CUSTOMER_NUMBER_COLUMN = "Customer Number"

# Cigna sheet column names
CIGNA_EMPLOYEE_ID_COLUMN = "Employee ID"
CIGNA_EMPLOYEE_NAME_COLUMN = "Employee Name"
CIGNA_MEDICAL_UNPOOLED_COLUMN = "Medical Unpooled"
CIGNA_MEDICAL_POOLED_COLUMN = "Medical Pooled/Fees"
CIGNA_DENTAL_COLUMN = "Dental"
CIGNA_VISION_COLUMN = "Vision"


# Cigna ID sheet column names
CIGNA_ID_MEMBER_ID_COLUMN = "Member ID"
CIGNA_ID_MEMBER_SSN_COLUMN = "Member SSN"
CIGNA_ID_RELATIONSHOP_COLUMN = "Relationship"

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
    def __init__(self, adp_file_full_path : str, insurance_file_full_path : str, id_file_full_path: str,  insurance_provider_type : INSURANCE_FORMAT_ENUM, plan_type : PLAN_TYPE_ENUM, output_folder : str, logger : Logger = None):
        self.adp_file_path = adp_file_full_path
        self.insurance_file_path = insurance_file_full_path
        self.id_file_full_path = id_file_full_path
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
        report_df = self._get_status_report(self.adp_file_path, self.insurance_file_path, self.id_file_full_path, self.insurance_provider_type, self.plan_type)
        if report_df is None:
            self._log_error(f"{self.plan_type.get_string()} with {self.insurance_provider_type.get_string()} is not supported.")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file_name = f"StatusReport_{self.insurance_provider_type.get_string()}_{self.plan_type.get_string()}_{timestamp}.xlsx"
        output_file_full_path = os.path.join(Path(self.output_folder), output_file_name)
        self._create_excel_file(report_df, output_file_full_path)

        if __debug__:
            csv_file_path = output_file_full_path.replace(".xlsx", '.csv')
            report_df.to_csv(csv_file_path, index=False)

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
        output_file_full_name = QUrl.fromLocalFile(output_file_full_name)
        output_file_sentence = f'Here is a <a href="{output_file_full_name.toString()}">{output_file_full_name.toString()}</a> word.'
        self._log_info(output_file_sentence)

    def _get_status_report(self, adp_file_full_path : str, insurance_file_full_path : str, id_file_full_path: str, insurance_format : INSURANCE_FORMAT_ENUM, plan_type : PLAN_TYPE_ENUM):
        if insurance_format == INSURANCE_FORMAT_ENUM.CIGNA:
            return self._get_status_report_for_cigna(adp_file_full_path, insurance_file_full_path, id_file_full_path)
        else:
            if plan_type == PLAN_TYPE_ENUM.DENTAL:
                return None
            elif plan_type == PLAN_TYPE_ENUM.EMPLOYEE_LIFE:
                return self._get_status_report_for_employee_life(adp_file_full_path, insurance_file_full_path, insurance_format)
            elif plan_type == PLAN_TYPE_ENUM.MEDICAL:
                return 
            elif plan_type == PLAN_TYPE_ENUM.VISION:
                return None
        
    def _get_status_report_for_cigna(self, adp_file_full_path : str, insurance_file_full_path : str, id_file_full_path : str) -> pd.DataFrame:
        adp_df = pd.read_excel(adp_file_full_path, header=None) # Assuming only one sheet in the excel file
        header_row_index = adp_df[adp_df.iloc[:, 5] == ADP_HIRE_DATE_COLUMN].index[0] # Find the header row (where column F has "HIRE DATE")
        adp_df = pd.read_excel(adp_file_full_path, header=header_row_index)       
        adp_df = adp_df[adp_df[ADP_PROVIDER_COLUMN] == INSURANCE_FORMAT_ENUM.CIGNA.get_string()] # should use a different enum that is dedicated for PROVIDER column
        #adp_filtered_df = adp_filtered_df.dropna(subset=[ADP_PLAN_TYPE_COLUMN]) # remove all rows with empty plan_type
        adp_filtered_df = adp_df[[ADP_NAME_COLUMN, ADP_TAX_ID_COLUMN, ADP_PLAN_TYPE_COLUMN, ADP_COVERAGE_LEVEL_VALUE_COLUMN]]
        adp_filtered_df[ADP_TAX_ID_COLUMN] = adp_filtered_df[ADP_TAX_ID_COLUMN].apply(self.keep_numbers_only)
        adp_filtered_df = adp_filtered_df.rename(columns={ADP_TAX_ID_COLUMN : CIGNA_ID_MEMBER_SSN_COLUMN})

        cigna_sheet_name = "Billing_Detail"
        insurance_df = pd.read_excel(insurance_file_full_path, cigna_sheet_name, header=11)
        index_of_endline = insurance_df.index[insurance_df[CIGNA_EMPLOYEE_NAME_COLUMN] == "Totals:"]
        insurance_df = insurance_df[[CIGNA_EMPLOYEE_ID_COLUMN, CIGNA_MEDICAL_POOLED_COLUMN, CIGNA_MEDICAL_UNPOOLED_COLUMN, CIGNA_DENTAL_COLUMN, CIGNA_VISION_COLUMN]]
        if not index_of_endline.empty:
            insurance_df = insurance_df.loc[:index_of_endline[0]-1] # keep everything before that row

        cigna_id_sheet_name = "Eligibility Roster Detail"
        insurance_id_df = pd .read_excel(id_file_full_path, cigna_id_sheet_name)
        insurance_id_df = insurance_id_df[[CIGNA_ID_MEMBER_ID_COLUMN, CIGNA_ID_MEMBER_SSN_COLUMN, CIGNA_ID_RELATIONSHOP_COLUMN]]
        insurance_id_df = insurance_id_df.rename(columns={CIGNA_ID_MEMBER_ID_COLUMN : CIGNA_EMPLOYEE_ID_COLUMN})
        insurance_id_df = insurance_id_df.dropna(subset=[CIGNA_ID_MEMBER_SSN_COLUMN]) # remove all rows with empty SSN
        insurance_id_df[CIGNA_ID_MEMBER_SSN_COLUMN] = insurance_id_df[CIGNA_ID_MEMBER_SSN_COLUMN].apply(self.keep_numbers_only)

        insurance_merged_df = pd.merge(insurance_df, insurance_id_df, on=[CIGNA_EMPLOYEE_ID_COLUMN], how="outer", suffixes=["_noID", "_ID"])
        
        merged_df = pd.merge(adp_filtered_df, insurance_merged_df, on=[CIGNA_ID_MEMBER_SSN_COLUMN], how="outer", suffixes=["_ADP", "_Cigna"])

        medical_status_column = "MEDICAL_STATUS"
        dental_status_column = "DENTAL_STATUS"
        vision_status_column = "VISION_STATUS"
        merged_df[COMMENT_COLUMN] = ""
        merged_df[medical_status_column] = ""
        merged_df[dental_status_column] = ""
        merged_df[vision_status_column] = ""
        new_comment = ""
        for row_index, row in merged_df.iterrows():
            status_column = COMMENT_COLUMN
            plan_type = row[ADP_PLAN_TYPE_COLUMN]
            if pd.isna(plan_type): # could be SP or CH, or could be missing employee in ADP
                relationship = row[CIGNA_ID_RELATIONSHOP_COLUMN]
                if relationship == CIGNA_ID_RELATIONSHIP_ENUM.EE.get_string():
                    new_comment = MATCHING_STATUS_ENUM.NEED_TO_BE_IN_ADP.get_string()
                elif relationship == CIGNA_ID_RELATIONSHIP_ENUM.SP.get_string() or relationship == CIGNA_ID_RELATIONSHIP_ENUM.CH.get_string():
                    new_comment = f"Employee family member: {relationship}"
            else:
                new_comment = f"{plan_type}: "
                if plan_type == PLAN_TYPE_ENUM.DENTAL.get_string():
                    status_column = dental_status_column
                    if row[CIGNA_DENTAL_COLUMN] > 0:
                        new_comment += MATCHING_STATUS_ENUM.GOOD_MATCHING.get_string()
                    else:
                        new_comment += f"Cigna should have bill greater than 0."
                elif plan_type == PLAN_TYPE_ENUM.MEDICAL.get_string():
                    status_column = medical_status_column
                    if row[CIGNA_MEDICAL_POOLED_COLUMN] + row[CIGNA_MEDICAL_UNPOOLED_COLUMN] > 0:
                        new_comment += MATCHING_STATUS_ENUM.GOOD_MATCHING.get_string()
                    else:
                        new_comment += f"Cigna should have bill greater than 0."
                elif plan_type == PLAN_TYPE_ENUM.VISION.get_string():
                    status_column = vision_status_column
                    if row[CIGNA_VISION_COLUMN] > 0:
                        new_comment += MATCHING_STATUS_ENUM.GOOD_MATCHING.get_string()
                    else:
                        new_comment += f"Cigna should have bill greater than 0."
                else:
                    self._log_error(f"Unsupported plan type {plan_type} from Cigna for {row[ADP_NAME_COLUMN]}, SSN: {row[CIGNA_ID_MEMBER_SSN_COLUMN]}")
                    new_comment += "Error"
            merged_df.at[row_index, status_column] = new_comment

        print(merged_df)
        merged_df = merged_df.groupby(CIGNA_ID_MEMBER_SSN_COLUMN, as_index=False).agg({
            ADP_NAME_COLUMN : "first",
            CIGNA_EMPLOYEE_ID_COLUMN: "first",
            CIGNA_MEDICAL_POOLED_COLUMN: "first",
            CIGNA_MEDICAL_UNPOOLED_COLUMN: "first",
            CIGNA_DENTAL_COLUMN: "first",
            CIGNA_VISION_COLUMN: "first",
            COMMENT_COLUMN : lambda x: " ".join(x),
            medical_status_column : lambda x: " ".join(x),
            dental_status_column : lambda x: " ".join(x),
            vision_status_column : lambda x: " ".join(x)
        })
        print(merged_df)


        return merged_df

    
    def _get_status_report_for_employee_life(self, adp_file_full_path: str, insurance_file_full_path: str, insurance_format: INSURANCE_FORMAT_ENUM) -> pd.DataFrame:
        adp_df = pd.read_excel(adp_file_full_path, header=None) # Assuming only one sheet in the excel file
        header_row_index = adp_df[adp_df.iloc[:, 5] == ADP_HIRE_DATE_COLUMN].index[0] # Find the header row (where column F has "HIRE DATE")
        adp_df = pd.read_excel(adp_file_full_path, header=header_row_index) 

        company_code = insurance_format.get_company_code_string()
        adp_filtered_df = adp_df[adp_df[ADP_COMPANY_CODE_COLUMN].isin(company_code) & (adp_df[ADP_PLAN_TYPE_COLUMN] == PLAN_TYPE_ENUM.EMPLOYEE_LIFE.get_string())]
        adp_filtered_df = adp_filtered_df[[ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN, ADP_HIRE_DATE_COLUMN, ADP_TERMINATION_DATE_COLUMN, ADP_ENROLLMENT_STATUS_COLUMN, ADP_EMPLOYEE_STATUS_COLUMN]] # strip the table and left only the wanted columns
        adp_filtered_df[ADP_DATE_OF_BIRTH_COLUMN] = adp_filtered_df[ADP_DATE_OF_BIRTH_COLUMN].apply(self._get_excel_serial_date)
        adp_filtered_df[ADP_HIRE_DATE_COLUMN] = adp_filtered_df[ADP_HIRE_DATE_COLUMN].apply(self._get_excel_serial_date)
        adp_filtered_df[ADP_TERMINATION_DATE_COLUMN] = adp_filtered_df[ADP_TERMINATION_DATE_COLUMN].apply(self._get_excel_serial_date)

        insurance_df = pd.read_excel(insurance_file_full_path, header=None) # Assuming only one sheet in the excel file
        header_row_index = insurance_df[insurance_df.iloc[:, 2] == EMPLOYEE_LIFE_FIRST_NAME_COLUMN].index[0] # Find the header row (where column C has "First Name")
        insurance_df = pd.read_excel(insurance_file_full_path, header=header_row_index)
        insurance_df[ADP_NAME_COLUMN] = insurance_df[EMPLOYEE_LIFE_LAST_NAME_COLUMN] + ", " + insurance_df[EMPLOYEE_LIFE_FIRST_NAME_COLUMN]
        insurance_df = insurance_df.rename(columns={EMPLOYEE_LIFE_DATE_OF_BIRTH_COLUMN : ADP_DATE_OF_BIRTH_COLUMN,
                                                    EMPLOYEE_LIFE_DATE_OF_HIRE_COLUMN : ADP_HIRE_DATE_COLUMN,
                                                    EMPLOYEE_LIFE_TERMINATION_DATE_COLUMN : ADP_TERMINATION_DATE_COLUMN})
        insurance_df = insurance_df[[ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN, ADP_HIRE_DATE_COLUMN, ADP_TERMINATION_DATE_COLUMN, EMPLOYEE_CUSTOMER_NUMBER_COLUMN]]     # strip the table and left only the wanted columns

        merged_df = pd.merge(insurance_df, adp_filtered_df, on=[ADP_NAME_COLUMN, ADP_DATE_OF_BIRTH_COLUMN], how="outer", suffixes=["_insurance", "_ADP"])
        merged_df[COMMENT_COLUMN] = ""
        new_insurance_date_of_hire_column = ADP_HIRE_DATE_COLUMN + "_insurance"
        new_insurance_termination_date_column = ADP_TERMINATION_DATE_COLUMN + "_ADP"
        new_adp_date_of_hire_column = ADP_HIRE_DATE_COLUMN + "_ADP"
        new_adp_termination_date_column = ADP_HIRE_DATE_COLUMN + "_ADP"

        new_comment = ""
        indices_to_pop = []
        for row_index, row in merged_df.iterrows():
            if row[COMMENT_COLUMN] != "":
                new_comment = MATCHING_STATUS_ENUM.DUPLICATE_FOUND.get_string()
            elif pd.isna(row[EMPLOYEE_CUSTOMER_NUMBER_COLUMN]):
                if row[ADP_EMPLOYEE_STATUS_COLUMN] == EMPLOYEE_STATUS_ENUM.LEAVE.get_string() and row[ADP_ENROLLMENT_STATUS_COLUMN] == ENROLLMENT_STATUS_ENUM.INACTIVE.get_string():
                    # it is supposed to be missing in insurance
                    indices_to_pop.append(row_index)
                else:
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
            merged_df.at[row_index, COMMENT_COLUMN] = new_comment

        merged_df = merged_df.drop(indices_to_pop)

        return merged_df
        
    def _get_last_and_first_name(self, full_name: str) -> tuple[str, str]:
        """
        full name is in format: "Last, First"
        Return: [Last, First]
        """
        parts = full_name.split(",")
        last_name = parts[0].strip()
        first_name = parts[1].strip()
        return last_name, first_name
    
    def _get_excel_serial_date(self, date_string, format: str = "%m/%d/%Y"):
        if not isinstance(date_string, str):
            if isinstance(date_string, (float, int)):
                return date_string
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
    
    def keep_numbers_only(self, num) -> str:
        if isinstance(num, str):
            num_string = re.sub(r"\D", "", num) # \D = non-digit
            num_string = num_string.zfill(9)
            return num_string
        elif isinstance(num, float):
            num_string = str(int(num))
            num_string = num_string.zfill(9)
            return num_string
        return None



