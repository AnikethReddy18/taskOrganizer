from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QGridLayout, QLabel, QWidget
import openpyxl

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Task Organizer")
        
        sheet = self.open_file()
        data = self.get_data(sheet)
        
        layout = QGridLayout()

        for row_index,row in enumerate(data):
            for column_index,cell in enumerate(row):
                layout.addWidget(QLabel(cell), row_index, column_index)

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

    def open_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Open Excel file", "", "Excel Files (*.xlsx)")
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

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


app = QApplication([])
window = MainWindow()
window.show()
app.exec()