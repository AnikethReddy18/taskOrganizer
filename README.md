# 📋 Task Organizer

**Task Organizer** is a simple desktop application built with **PyQt6** and **OpenPyXL** that helps you manage tasks stored in an Excel `.xlsx` file. It reads tasks from the spreadsheet, displays incomplete ones in a scrollable GUI, lets you mark them as **Done**, copy associated links, and even randomly pick a task to focus on!

---

## 🔧 Features

- 📂 Reads tasks from a local Excel file.
- ✅ Marks tasks as **Done** and updates the Excel file.
- 🔗 Allows copying task-related links to clipboard.
- 🎲 Randomly selects one of the pending tasks and highlights it.
- 🔄 Automatically hides completed tasks from the list.

---

## 📁 Excel File Format

The app expects an Excel file structured like this:

| Task Title | Description (can contain hyperlink) | 
|------------|-------------------------------------|
| Task 1     | Description 
| Task 2     |Description                 |  

- The app adds the **"Task Status"** column automatically .
- Tasks marked as **"Done"** are hidden in the UI.
- Descriptions with hyperlinks can be copied via the **Copy** button.

---

### 🔧 Installation

#### 📦 Requirements

- Python 3.7+
- PyQt6
- openpyxl

1. Clone this repository:
   ```bash
   git clone https://github.com/AnikethReddy18/taskOrganizer
   cd taskOrganizer
   ```

 2. Install Dependencies
    ```bash
    pip install pyqt6 openpyxl
    ```

 

### 🛠️ To-Do
- Highlight completed tasks with a different style instead of hiding.

- Add search/filter options.

   
