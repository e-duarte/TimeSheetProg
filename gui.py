import sys
from PySide6.QtWidgets import (QLineEdit, QPushButton, QApplication,
    QVBoxLayout, QDialog, QHBoxLayout, QWidget, QLabel, QMainWindow, QComboBox)
from PySide6 import QtWidgets
from calendar_util import CalendarUtil
from employee import EmployeesLoader
from timesheet_book import Book

class MonthPicker(QWidget):
    def __init__(self, list, parent=None):
        super().__init__(parent)
        self.month = ''
        self.label = QLabel('Escolha o mês:')
        self.selector = QComboBox()
        self.selector.addItems(list)

        layout = QHBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.selector)

        self.month = self.selector.currentText()

        self.selector.currentTextChanged.connect(self.getChonsenMonth)

        self.setLayout(layout)

    def getChonsenMonth(self, s):
        print(s)
        self.month = s


class InputPath(QWidget):
    def __init__(self, button_title, type,parent=None):
        super().__init__(parent)
        self.path = ''
        layout = QHBoxLayout()
        self.edit = QLineEdit()

        self.button = QPushButton(button_title)
        layout.addWidget(self.edit)
        layout.addWidget(self.button)

        if type == 'file':
            self.button.clicked.connect(self.openFileManager)
        else:
            self.button.clicked.connect(self.openDirectoryManager)

        self.setLayout(layout)

    def openFileManager(self):
        filename = str(QtWidgets.QFileDialog.getOpenFileName(self, 'Abrir planilha', './', "Docx files (*.xlsx)" )[0])
        self.path = filename
        self.edit.setText(filename)

    def openDirectoryManager(self):
        directory = str(
            QtWidgets.QFileDialog.getSaveFileName(self, 'Salvar arquivo docx', './', "Docx files (*.docx)")[0]
        )
        directory = f'{directory}.docx'
        self.path = directory
        self.edit.setText(directory)


class BookDoneDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout()
        message = QLabel('O Livro de ponto foi criado')
        self.layout.addWidget(message)
        self.setLayout(self.layout)

class MainScreen(QMainWindow):    
    def __init__(self, parent=None):
        super(MainScreen, self).__init__(parent)
        self.setWindowTitle('Livro de Ponto - TimeSheetProg')
        container = QWidget()
        self.setCentralWidget(container)
        layout = QVBoxLayout()
        container.setLayout(layout)

        self.fp_label = QLabel('Planilha')
        self.filePathWidget = InputPath('...', 'file')
        
        self.selector = MonthPicker(CalendarUtil().months)

        self.doc_label = QLabel('Salvar Livro de Ponto')
        self.docPathWidget = InputPath('...', 'directory')

        self.button = QPushButton('Gerar Livro')
        self.button.clicked.connect(self.buildBook)

        layout.addWidget(self.fp_label)
        layout.addWidget(self.filePathWidget)
        layout.addWidget(self.selector)
        layout.addWidget(self.doc_label)
        layout.addWidget(self.docPathWidget)
        layout.addWidget(self.button)

        self.setLayout(layout)

    def buildBook(self):
        path_sheet = self.filePathWidget.path
        month = self.selector.month
        path_docx = self.docPathWidget.path

        print('Planilha escolhida: ', path_sheet)
        print('Mês:', month)
        print('Destino', path_docx)

        loader = EmployeesLoader(path_sheet)
        employees = loader.df2Dict()    
        book = Book(employees,month)
        book.build_pages()
        book.save(path_docx)
        print('Book was build')

        dlg = BookDoneDialog(self)
        dlg.setWindowTitle("Concluido!")
        dlg.exec()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = MainScreen()
    main.show()
    sys.exit(app.exec())