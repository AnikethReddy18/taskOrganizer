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
        #file_path, _ = file_dialog.getOpenFileName(self, "Open Excel file", "", "Excel Files (*.xlsx)")
        workbook = openpyxl.load_workbook("/home/aniketh/Downloads/DSA.xlsx")
        sheet = workbook.active
 
        if "done" in workbook.sheetnames:
            pass
        else:
            workbook.create_sheet('done')
            workbook.save("/home/aniketh/Downloads/DSA.xlsx")
        return sheet

    def get_data(self, sheet):
        data = []

        for row in sheet:
            if row[0].value == None:
                break
            title = row[0].value
            description = row[1].value
            link = row[1].hyperlink.target if row[1].hyperlink else None

            data.append([title, description, link])
        
        return data

    def display_data(self, data):
        layout = QGridLayout()

        for row_index,row in enumerate(data):
            for column_index,cell in enumerate(row):
                label = QLabel(cell)
                label.setWordWrap(True)
                layout.addWidget(label, row_index, column_index)
            layout.addWidget(QPushButton("Done"), row_index, len(row))

        widget = QWidget()
        widget.setLayout(layout)
        scroll = QScrollArea()
        scroll.setWidget(widget)
        self.setCentralWidget(scroll)

    def 

app = QApplication([])
window = MainWindow()
window.show()
app.exec()