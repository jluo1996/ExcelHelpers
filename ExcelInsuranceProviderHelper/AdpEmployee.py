from EmployeeBase import EmployeeBase
from InsuranceStatusHelperEnum import PLAN_TYPE_ENUM, INSURANCE_PROVIDER_ENUM


class AdpEmployee(EmployeeBase):
    def __init__(self, full_name : str, date_of_birth, hire_date, termination_date, plans : list[PLAN_TYPE_ENUM], insurance_provider : INSURANCE_PROVIDER_ENUM):
        first_name, last_name = self.get_first_last_name(full_name)
        super().__init__(first_name, last_name, date_of_birth)
        self.hire_date = hire_date
        self.termination_date = termination_date
        self.plans = plans
        self.insurance_provider = insurance_provider

    def get_first_last_name(self, full_name : str) -> tuple[str, str]:
        parts = full_name.split(",")
        last_name = parts[0].strip()
        first_name = parts[1].strip()
        return first_name, last_name
    