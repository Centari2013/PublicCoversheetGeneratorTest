import sqlite3
import os
import sys
from PyQt6.QtWidgets import *
from PyQt6.QtCore import QRegularExpression, QObject, pyqtSignal, QThread
from PyQt6.QtGui import QRegularExpressionValidator
from docxtpl import DocxTemplate
import time


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class ccGenWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.template = DocxTemplate(resource_path("CC PROCESS TEMPLATE.docx"))
        self.db = sqlite3.connect(resource_path("clients.db"))
        self.cur = self.db.cursor()

        self.setWindowTitle("CC Coversheet Generator")
        self.setFixedSize(274,358)

        layout = QVBoxLayout()

        client_label = QLabel("Select Client: ")
        self.client_select = QComboBox()
        self.populateClientSelect()

        amount_label = QLabel("Enter Processed Amount: ")

        self.amount_input = QLineEdit()
        reg_ex = QRegularExpression("[0-9]+.?[0-9]{0,2}")
        reg_ex_validator = QRegularExpressionValidator(reg_ex)
        self.amount_input.setValidator(reg_ex_validator)
        self.amount_input.returnPressed.connect(self.generateDoc)
        self.amount_input.returnPressed.connect(self.sleepInput)

        self.amount_error = QLabel("Please enter an amount.")
        self.amount_error.setStyleSheet("color: red")
        self.amount_error.hide()

        self.date_picker = QCalendarWidget()

        self.submit_button = QPushButton("Generate Doc")
        self.submit_button.clicked.connect(self.generateDoc)
        self.submit_button.clicked.connect(self.sleepInput)

        widgets = [client_label, self.client_select, amount_label, self.amount_input, self.amount_error,
                   self.date_picker, self.submit_button]

        for w in widgets:
            layout.addWidget(w)
        widget = QWidget()
        widget.setLayout(layout)

        # Set the central widget of the Window. Widget will expand
        # to take up all the space in the window by default.
        self.setCentralWidget(widget)


    def populateClientSelect(self):
        tups = self.cur.execute("""SELECT name FROM client
                                ORDER BY name ASC;""")

        clients = [name for t in tups for name in t]
        self.client_select.addItems(clients)

    def generateDoc(self):
        if self.amount_input.text() == "":
            self.amount_error.show()
        else:
            self.amount_error.hide()
            self.submit_button.setEnabled(False)
            self.amount_input.setEnabled(False)

            client_name = self.client_select.currentText()
            client_code = self.cur.execute("SELECT code FROM client WHERE name = ?", (client_name,))
            client_code = [code for t in client_code for code in t][0]

            processed_amt = float(self.amount_input.text())
            processed_amt = '{:,.2f}'.format(float(processed_amt))

            date = self.date_picker.selectedDate().toPyDate().strftime("%m-%d-%Y")

            fill_in = {"client": client_name,
                       "code": client_code,
                       "amount": processed_amt,
                       "date": date
                       }

            self.template.render(fill_in)
            self.template.save(resource_path("coversheet.docx"))


            self.amount_input.clear()

            os.startfile(resource_path("coversheet.docx"), "print")


    class Worker(QObject):
        finished = pyqtSignal()

        def run(self):
            """Long-running task."""
            time.sleep(3)
            self.finished.emit()

    def sleepInput(self):
        # Step 2: Create a QThread object
        self.thread = QThread()
        # Step 3: Create a worker object
        self.worker = self.Worker()
        # Step 4: Move worker to the thread
        self.worker.moveToThread(self.thread)
        # Step 5: Connect signals and slots
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        # Step 6: Start the thread
        self.thread.start()

        # Final resets
        self.submit_button.setEnabled(False)
        self.amount_input.setEnabled(False)

        self.thread.finished.connect(lambda: self.submit_button.setEnabled(True) )
        self.thread.finished.connect(lambda: self.amount_input.setEnabled(True))



def main():
    app = QApplication(sys.argv)
    window = ccGenWindow()
    window.show()

    app.exec()


if __name__ == "__main__":
    main()
