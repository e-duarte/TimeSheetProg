import pandas as pd

class EmployeesLoader:
    def __init__(self, path):
        self.employees = pd.read_excel(path, dtype={'Telefone': object, 'Turnos': str})
        self.employees.fillna('', inplace=True)

    def df2Dict(self):
        dict_employees = []
        columns = self.employees.columns.to_list()
        
        for a in self.employees.itertuples():
            vals = list(a)[1:]
            dict_employees.append(
                {
                    columns[0]: vals[0],
                    columns[1]: vals[1],
                    columns[2]: vals[2],
                    columns[3]: vals[3],
                    columns[4]: vals[4],
                    columns[5]: vals[5],
                    columns[6]: str(vals[6]).split('-')
                }
            )
        
        return dict_employees
