from PyQt6.QtWidgets import (QApplication, QMainWindow, QFileDialog, QScrollArea,
                             QGridLayout, QLabel, QWidget, QPushButton)

import openpyxl

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Task Organizer")
        self.resize(820, 400)

        self.undone_rows =[]
        self.done_rows = []
        sheet = self.open_file()
        self.get_data(sheet)

        self.display_data()

    def open_file(self):
        #file_dialog = QFileDialog()
        #self.file_path, _ = file_dialog.getOpenFileName(self, "Open Excel file", "", "Excel Files (*.xlsx)")
        self.file_path = "/home/aniketh/Downloads/DSA.xlsx"
        self.workbook = openpyxl.load_workbook(self.file_path)
        sheet = self.workbook.active
 
        if "done" in self.workbook.sheetnames:
            pass
        else:
            self.workbook.create_sheet('done')
        return sheet

    def get_data(self, sheet):

        self.done_status = []
        for i, row in enumerate(sheet.iter_rows()):
            if i == 0:
                continue
            if row[0].value == None:
                break
            title = row[0].value
            description = row[1].value
            link = row[1].hyperlink.target if row[1].hyperlink else None
            if row[2].value == "Done":
                self.done_rows.append([title, description, link])
            else:
                self.undone_rows.append([title, description, link])        
    
    def display_data(self):
        layout = QGridLayout()

        for column_index,cell in enumerate(list(self.workbook.active.iter_rows())[0]):
                label = QLabel(cell.value)
                label.setWordWrap(True)
                layout.addWidget(label, 0, column_index)

        select_random_button = QPushButton("Select Random")
        layout.addWidget(select_random_button, 0, 2, 1, 2)    
        
        for row_index,row in enumerate(self.undone_rows):
            for column_index,cell in enumerate(row):
                label = QLabel(cell)
                label.setWordWrap(True)
                layout.addWidget(label, row_index+1, column_index)
            button = QPushButton("Done")
            button.clicked.connect(self.done_button_was_pressed)
            button.setProperty("index", row_index)
            self.columns = len(row)
            
            layout.addWidget(button, row_index+1, self.columns)

        widget = QWidget()
        widget.setLayout(layout)
        scroll = QScrollArea()
        scroll.setWidget(widget)
        self.setCentralWidget(scroll)

    def done_button_was_pressed(self):
        index = self.sender().property('index')
        self.workbook.active.cell(row=index+1, column=self.columns).value = "Done"
        self.workbook.save(self.file_path)

        removed_row = self.undone_rows.pop(index)
        self.done_rows.append(removed_row)
        self.setCentralWidget(QWidget())
        self.display_data()
    
    def select_random_button_was_pressed(self):
        pass 

app = QApplication([])
window = MainWindow()
window.show()
app.exec()