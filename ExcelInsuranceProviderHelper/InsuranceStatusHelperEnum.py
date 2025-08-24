
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
class INSURANCE_PROVIDER_ENUM(Enum):
    CIGNA = 0
    BSS = 1
    BFS = 2

    def get_string(self):
        return {
            INSURANCE_PROVIDER_ENUM.CIGNA : "Cigna",
            INSURANCE_PROVIDER_ENUM.BSS : "bss",
            INSURANCE_PROVIDER_ENUM.BFS : "bfs"
        }[self]

# Define an enumeration for different matching statuses
class MATCHING_STATUS_ENUM(Enum):
    GOOD_MATCHING = 0
    DUPLICATE_FOUND = 1
    MISMATCHING_START_DATE = 2
    MISMATCHING_END_DATE = 3
    NOT_EXIST = 4

    def get_string(self):
        return {
            MATCHING_STATUS_ENUM.GOOD_MATCHING: "Good Matching",
            MATCHING_STATUS_ENUM.DUPLICATE_FOUND: "Duplicate Found",
            MATCHING_STATUS_ENUM.MISMATCHING_START_DATE: "Mismatching Start Date",
            MATCHING_STATUS_ENUM.MISMATCHING_END_DATE: "Mismatching End Date",
            MATCHING_STATUS_ENUM.NOT_EXIST: "Not Exist"
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

