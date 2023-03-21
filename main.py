from employee import EmployeesLoader
from timesheet_book import Book
import tkinter as tk
from tkinter.filedialog import askopenfilename


if __name__ == "__main__":
    fn = askopenfilename()
    print("user chose", fn)
    loader = EmployeesLoader(fn)
    employees = loader.df2Dict()    
    book = Book(employees,)
    book.build_pages()
    book.save()
    print('Ending TimeSheetProg')