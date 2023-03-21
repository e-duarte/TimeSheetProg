import sys
from tkinter.filedialog import askopenfilename
from PySide6.QtWidgets import (QLineEdit, QPushButton, QApplication,
    QVBoxLayout, QDialog, QHBoxLayout)

class MainScreen(QDialog):

    def __init__(self, parent=None):
        super(MainScreen, self).__init__(parent)
        # Create widgets
        self.edit = QLineEdit()
        self.button = QPushButton("Abrir")
        # Create layout and add widgets
        layout = QHBoxLayout()
        layout.addWidget(self.edit)
        layout.addWidget(self.button)
        # Set dialog layout
        self.setLayout(layout)
        # Add button signal to greetings slot
        self.button.clicked.connect(self.greetings)

    # Greets the user
    def greetings(self):
        fn = askopenfilename()
        self.edit.setText(fn)
        print(f"user chose {self.edit.text()}")

if __name__ == '__main__':
    # Create the Qt Application
    app = QApplication(sys.argv)
    # Create and show the form
    main = MainScreen()
    main.show()
    # Run the main Qt loop
    sys.exit(app.exec())