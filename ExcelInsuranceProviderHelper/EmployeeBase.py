class EmployeeBase(object):
    def __init__(self, first_name : str, last_name : str, date_of_birth):
        self.first_name = first_name
        self.last_name = last_name
        self.date_of_birth = date_of_birth

    def get_employee_id(self):
        return f"{self.first_name}{self.last_name}{self.date_of_birth}"
