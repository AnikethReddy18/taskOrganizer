from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog
import openpyxl

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Task Organizer")
        self.open_file()

    def open_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Open Excel file", "", "Excel Files (*.xlsx)")
        workbook = openpyxl.load_workbook(filename=file_path)
        sheet = workbook.active
        things = []
        for row in sheet:
            if row[0].value == None:
                break
            title = row[0].value
            description = row[1].value
            link = row[1].hyperlink.target if row[1].hyperlink else None

            things.append([title, description, link])
        print(things)


app = QApplication([])
window = MainWindow()
window.show()
app.exec()