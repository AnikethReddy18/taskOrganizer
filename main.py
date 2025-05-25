from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Task Organizer")
        self.open_file()

    def open_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Open Excel file", "", "Excel Files (*.xlsx)")



app = QApplication([])
window = MainWindow()
window.show()
app.exec()