
from enum import Enum


# Define an enumeration for different plan types
class PLAN_TYPE_ENUM(Enum):
    DENTAL = "Dental"
    EMPLOYEE_LIFE = "Employee Life"
    MEDICAL = "Medical"
    VISION = "Vision"

# Define an enumeration for different insurance providers
class INSURANCE_PROVIDER_ENUM(Enum):
    CIGNA = "Cigna"
    BSS = "BSS" # TODO: find the full name
    BFS = "BFS" # TODO: find the full name

# Define an enumeration for different matching statuses
class MATCHING_STATUS_ENUM(Enum):
    GOOD_MATCHING = "Good Matching"
    DUPLICATE_FOUND = "Duplicate Found"
    MISMATCHING_START_DATE = "Mismatching Start Date"
    MISMATCHING_END_DATE = "Mismatching End Date"
    NOT_EXIST = "Not Exist"


# region ADP enums

# Define an enumeration for different insurance enrollment statuses in ADP
class ENROLLMENT_STATUS_ENUM(Enum):
    ACTIVE = "Active"
    INACTIVE = "Inactive"

# endregion

