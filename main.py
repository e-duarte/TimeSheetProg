import pandas as pd
from timesheet_book import Book
import tkinter as tk
from tkinter.filedialog import askopenfilename

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

if __name__ == "__main__":
    fn = askopenfilename()
    print("user chose", fn)
    loader = EmployeesLoader(fn)
    employees = loader.df2Dict()    
    book = Book(employees,)
    book.build_pages()
    book.save()
    print('Ending TimeSheetProg')