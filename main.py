from PyQt6.QtWidgets import (QApplication, QMainWindow, QFileDialog, QScrollArea,
                             QGridLayout, QLabel, QWidget, QPushButton)

import openpyxl

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Task Organizer")
        self.resize(820, 400)

        sheet = self.open_file()
        data = self.get_data(sheet)

        self.display_data(data)

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
        data = []

        self.done_status = []
        for row in sheet:
            if row[0].value == None:
                break
            title = row[0].value
            description = row[1].value
            link = row[1].hyperlink.target if row[1].hyperlink else None
            self.done_status.append(1 if row[2].value == "Done" else 0)

            data.append([title, description, link])
        return data
    
    def display_data(self, data):
        layout = QGridLayout()

        for row_index,row in enumerate(data):
            if(self.done_status[row_index] == 1):
                continue
            for column_index,cell in enumerate(row):
                label = QLabel(cell)
                label.setWordWrap(True)
                layout.addWidget(label, row_index, column_index)
            button = QPushButton("Done")
            button.clicked.connect(self.done_button_was_pressed)
            button.setProperty("index", row_index)
            self.columns = len(row)
            
            layout.addWidget(button, row_index, self.columns)

        widget = QWidget()
        widget.setLayout(layout)
        scroll = QScrollArea()
        scroll.setWidget(widget)
        self.setCentralWidget(scroll)

    def done_button_was_pressed(self):
        index = self.sender().property('index')
        self.workbook.active.cell(row=index+1, column=self.columns).value = "Done"
        self.workbook.save(self.file_path)

        data = self.get_data(self.workbook.active)
        self.setCentralWidget(QWidget())
        self.display_data(data)
    
        

app = QApplication([])
window = MainWindow()
window.show()
app.exec()