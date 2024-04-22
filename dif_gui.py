import os
import sys
from PySide2.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, \
    QListWidget, QMessageBox, QComboBox, QFileDialog
from PySide2.QtCore import Qt
import subprocess
import time
import re
import docx2txt
import PyPDF2
import pandas as pd


class File_search(QWidget):
    def __init__(self):
        super().__init__()
        self.show_Gui()

    def show_Gui(self):
        # Setup the main GUI layout and components
        self.setWindowTitle('Document Retrieval Search Software')
        self.setGeometry(100, 100, 800, 600)
        self.v_layout = QVBoxLayout(self)

        self.introduction = QLabel('''<h2 style="text-align: center;font-family: Arial">Document Retrieval Software Description </h2><br/>\
                                        <font style="font-size:20px;font-family: Arial">
                                        1.Select the directory to find the desire file from your directory.<br/>\
                                        2.Type the dire file name to search for. the file can be Word/Excel/Pdf/Text/Python.<br/>\
                                        3.File name should be .docx/.xlsx/.pdf/.txt/.py, so this software's finding ability is much more accurate.
                                        </font>''')
        self.introduction.setAlignment(Qt.AlignTop)
        self.v_layout.addWidget(self.introduction)

        self.setupDirectoryInput()
        self.setupFileTypeInput()
        self.setupKeywordInput()

        self.result_list = QListWidget()
        self.v_layout.addWidget(self.result_list)
        self.result_list.itemClicked.connect(self.open_file)

    def setupDirectoryInput(self):
        # Directory input setup
        widget = QWidget()
        layout = QHBoxLayout(widget)
        label = QLabel('<font style="font-size:18px;font-family: Arial;">Directory:</font>')
        self.directory_input = QLineEdit()
        self.directory_input.setPlaceholderText('Input directory...')
        browse_btn = QPushButton('Browse')
        browse_btn.clicked.connect(self.browse_directory)
        layout.addWidget(label)
        layout.addWidget(self.directory_input)
        layout.addWidget(browse_btn)
        self.v_layout.addWidget(widget)

    def setupFileTypeInput(self):
        # File type input setup
        widget = QWidget()
        layout = QHBoxLayout(widget)
        label = QLabel('<font style="font-size:18px;font-family: Arial;">File Type:</font>')
        self.file_type_combo = QComboBox()
        self.file_type_combo.addItems(["All Files", ".docx", ".xlsx", ".pdf", ".txt", ".py"])
        layout.addWidget(label)
        layout.addWidget(self.file_type_combo)
        self.v_layout.addWidget(widget)

    def setupKeywordInput(self):
        # Keyword input setup
        widget = QWidget()
        layout = QHBoxLayout(widget)
        label = QLabel('<font style="font-size:18px;font-family: Arial;">Keyword:</font>')
        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText('Enter keyword to search...')
        search_btn = QPushButton('Search')
        search_btn.clicked.connect(self.perform_keyword_search)
        layout.addWidget(label)
        layout.addWidget(self.keyword_input)
        layout.addWidget(search_btn)
        self.v_layout.addWidget(widget)

    def browse_directory(self):
        # Function to browse and select a directory
        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            self.directory_input.setText(directory)

    def perform_keyword_search(self):
        # Perform search based on input criteria
        directory = self.directory_input.text()
        keyword = self.keyword_input.text()
        file_extension = self.file_type_combo.currentText()

        if not directory:
            QMessageBox.warning(self, "Warning", "Please enter a directory.")
            return

        if not keyword:
            QMessageBox.warning(self, "Warning", "Please enter a keyword to search for.")
            return

        self.result_list.clear()
        self.search_files(directory, keyword, file_extension)

    def search_files(self, directory, keyword, file_extension):
        # Search files in the specified directory matching the file extension and containing the keyword
        matched_files = []
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file_extension == "All Files" or file.endswith(file_extension):
                    full_path = os.path.join(root, file)
                    if self.contains_keyword(full_path, keyword):
                        matched_files.append(full_path)

        if matched_files:
            self.result_list.addItems(matched_files)
        else:
            self.result_list.addItem("No matching files found.")

    def contains_keyword(self, file_path, keyword):
        # Check if the file contains the keyword
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                return keyword in file.read()
        except Exception as e:
            return False

    def open_file(self, item):
        # Function to open the selected file
        file_path = item.text()
        if sys.platform == "win32":
            os.startfile(file_path)
        else:
            subprocess.run(['open', file_path])


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = File_search()
    window.show()
    sys.exit(app.exec_())
