
from enum import Enum


# Define an enumeration for different plan types
class PLAN_TYPE_ENUM(Enum):
    DENTAL = 0
    EMPLOYEE_LIFE = 1
    MEDICAL = 2
    VISION = 3

    def get_string(self):
        return {
            PLAN_TYPE_ENUM.DENTAL : "Dental",
            PLAN_TYPE_ENUM.EMPLOYEE_LIFE : "Employee Life",
            PLAN_TYPE_ENUM.MEDICAL: "Medical",
            PLAN_TYPE_ENUM.VISION: "Vision" 
        }[self]

# Define an enumeration for different insurance providers
class INSURANCE_FORMAT_ENUM(Enum):
    CIGNA = 0
    BSS = 1
    BFS = 2

    def get_string(self):
        return {
            INSURANCE_FORMAT_ENUM.CIGNA : "Cigna",
            INSURANCE_FORMAT_ENUM.BSS : "bss",
            INSURANCE_FORMAT_ENUM.BFS : "bfs"
        }[self]
    
    def get_company_code_enum(self):
        """
        ReturnL: list of COMPANY_CODE_ENUM
        """
        match self:
            case INSURANCE_FORMAT_ENUM.BSS:
                return [COMPANY_CODE_ENUM.E30]
            case INSURANCE_FORMAT_ENUM.BFS:
                return [COMPANY_CODE_ENUM.E30, COMPANY_CODE_ENUM.E9Y]
            case _:
                return [COMPANY_CODE_ENUM.UNKNOWN]
            
    def get_company_code_string(self):
        """
        Return: list[str]
        """
        enums = self.get_company_code_enum()
        output = []
        for enum in enums:
            output.append(enum.get_string())
        return output
    
class EMPLOYEE_STATUS_ENUM(Enum):
    ACTIVE = 0
    TERMINATED = 1
    LEAVE = 2

    def get_string(self):
        return {
            EMPLOYEE_STATUS_ENUM.ACTIVE: "Active",
            EMPLOYEE_STATUS_ENUM.TERMINATED: "Terminated",
            EMPLOYEE_STATUS_ENUM.LEAVE: "Leave"
        }[self]
    
class COMPANY_CODE_ENUM(Enum):
    UNKNOWN = -1
    E30 = 0 # BFS
    E9V = 1 # BSS
    E9Y = 2 # BFS

    def get_string(self):
        return {
            COMPANY_CODE_ENUM.UNKNOWN: "Unknown",
            COMPANY_CODE_ENUM.E30: "E30",
            COMPANY_CODE_ENUM.E9V: "E9V",
            COMPANY_CODE_ENUM.E9Y: "E9Y"
        }[self]

# Define an enumeration for different matching statuses
class MATCHING_STATUS_ENUM(Enum):
    GOOD_MATCHING = 0
    DUPLICATE_FOUND = 1
    MISMATCHING_START_DATE = 2
    MISMATCHING_END_DATE = 3
    EXIST_ONLY_IN_ADP = 4
    NEED_TO_BE_IN_ADP = 5

    def get_string(self):
        return {
            MATCHING_STATUS_ENUM.GOOD_MATCHING: "Good Matching",
            MATCHING_STATUS_ENUM.DUPLICATE_FOUND: "Duplicate Found",
            MATCHING_STATUS_ENUM.MISMATCHING_START_DATE: "Mismatching Start Date",
            MATCHING_STATUS_ENUM.MISMATCHING_END_DATE: "Mismatching End Date",
            MATCHING_STATUS_ENUM.EXIST_ONLY_IN_ADP: "Exist only in ADP",
            MATCHING_STATUS_ENUM.NEED_TO_BE_IN_ADP: "Need to be in ADP"
        }[self]


# region ADP enums

# Define an enumeration for different insurance enrollment statuses in ADP
class ENROLLMENT_STATUS_ENUM(Enum):
    ACTIVE = 0
    INACTIVE = 1

    def get_string(self):
        return {
            ENROLLMENT_STATUS_ENUM.ACTIVE: "Active",
            ENROLLMENT_STATUS_ENUM.INACTIVE: "Inactive"
        }[self]

    

# endregion

