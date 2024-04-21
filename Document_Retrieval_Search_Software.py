    # Import all necessary libraries and their functions.
import os
from PySide2.QtWidgets import *
from PySide2.QtCore import *
import time
import re
import docx2txt
import PyPDF2
import pandas as pd

# custom widget class File_search
class File_search(QWidget):
    def __init__(self):
        # call the superclass
        super().__init__()
        # create a vertical layout
        self.show_Gui()
    #display user interface
    def show_Gui(self):
        #set the window display title
        self.setWindowTitle('Document Retrieval Search Software')
        # self.showMaximized()
        self.v_layout = QVBoxLayout()
        self.setLayout(self.v_layout)
        # document retrieval software description box
        self.introduction = QLabel()
        self.introduction.setText('''<h2 style="text-align: center;font-family: Arial">Document Retrieval Software Description </h2><br/>\
                                    <font style="font-size:20px;font-family: Arial">
                                    1.Select the directory to find the desire file from your directory.<br/>\
                                    2.Type the dire file name to search for. the file can be Word/Excel/Pdf/Text/Python.<br/>\
                                    3.File name should be .docx/.xlsx/.pdf/.txt/.py, so this software's finding ability is much more accurate.
                                    </font>''')
        self.introduction.setStyleSheet('border: 1px solid #666;line-height:3em;height:5000px')
        self.introduction.setAlignment(Qt.AlignTop)
        # create tooltip text, text box, and browse button
        self.widget1 = QWidget()
        self.layout1 = QHBoxLayout(self.widget1)
        self.tip_text1 = QLabel('<font style="font-size:18px;font-family: Arial;">Directory_:</font>')
        self.text1 = QLineEdit()
        self.text1.setPlaceholderText('Input directory...')
        self.text1.setMinimumHeight(30)
        self.text1.setMaximumHeight(30)
        self.browse_btn = QPushButton('Browse')
        self.browse_btn.setStyleSheet('font-size: 18px; font-family: Arial; ')
        self.browse_btn.clicked.connect(self.browse_begin)
        self.layout1.addWidget(self.tip_text1)
        self.layout1.addWidget(self.text1)
        self.layout1.addWidget(self.browse_btn)
        # create tooltip text, text box, and search button
        self.widget2 = QWidget()
        self.layout2 = QHBoxLayout(self.widget2)
        self.tip_text2 = QLabel('<font style="font-size:18px;font-family: Arial">Find name:</font>')
        self.text2 = QLineEdit()
        self.text2.setPlaceholderText('Input File Name to search...')
        self.text2.setMinimumHeight(30)
        self.text2.setMaximumHeight(30)
        self.search_btn = QPushButton('Search')
        self.search_btn.setStyleSheet('font-size: 18px; font-family: Arial; ')
        self.search_btn.clicked.connect(self.search_begin)
        self.layout2.addWidget(self.tip_text2)
        self.layout2.addWidget(self.text2)
        self.layout2.addWidget(self.search_btn)
        # create tooltip text, text box, and keyword search button
        self.widget3 = QWidget()
        self.layout3 = QHBoxLayout(self.widget3)
        self.tip_text3 = QLabel('<font style="font-size:18px;font-family: Arial">Keyword_:</font>')
        self.text3 = QLineEdit()
        self.text3.setPlaceholderText('Input Keyword to find...')
        self.text3.setMinimumHeight(30)
        self.text3.setMaximumHeight(30)
        self.find_btn = QPushButton('Find')
        self.find_btn.setStyleSheet('font-size: 18px; font-family: Arial; ')
        self.find_btn.clicked.connect(self.keyword_find_begin)
        self.layout3.addWidget(self.tip_text3)
        self.layout3.addWidget(self.text3)
        self.layout3.addWidget(self.find_btn)
        #display all layouts
        self.v_layout.addWidget(self.introduction, stretch=1)
        self.v_layout.addWidget(self.widget1)
        self.v_layout.addWidget(self.widget2)
        self.v_layout.addWidget(self.widget3)
        self.result_flag = 0
        self.key_word_result_flag = 0
        self.show()
    # user begins to browse the catalog
    def browse_begin(self):
        # open the directory dialog box, if the user clicks OK, the directory will be selected
        file_name = QFileDialog.getExistingDirectory(self, "Select Directory",)
        if file_name:
            #set the text in the text input box to the selected directory path by changing '/' to ''
            self.text1.setText(file_name.replace('/', '\\'))
    #perform file search operations in a directory and its subdirectories
    def search_begin(self):
        dirname = self.text1.text()
        filename = self.text2.text()
        self.allpathes = []
        # Check if both the directory path and filename are provided. If not, it will display a warning message using the QMessageBox.warning method and return from the function
        if not dirname and not filename:
            QMessageBox.warning(self, "Warning", "Please enter a Directory path and a File name.")
            return
        if not dirname :
            QMessageBox.warning(self, "Warning", "Please enter a Directory path.")
            return
        if not filename:
            QMessageBox.warning(self, "Warning", "Please enter a File name.")
            return
        self.search_begin_time = time.time()
        directory = self.text1.text()
        # If the specified directory path exists, then the method uses the os.walk function to traverse the directory and its subdirectories
        if os.path.exists(directory):
            for root, dirs, files in os.walk(directory):  
                # This method uses the re.search function to check if the filename meets the search criteria
                for file in files:
                    # If a match is found, use the os.path.join method to obtain the full path of the file and add it to the allpaths list
                    if re.search(filename.split('\\')[-1], file): 
                        path = os.path.join(root, file)  
                        self.allpathes.append(path)  
            self.show_search_result(self.allpathes)
        # If the specified directory path does not exist, then display a warning message
        else:
            QMessageBox.warning(self, 'Warning', 'Path not exist, Please input a correct path name')
    #complete the file search operation and display it in the window
    def show_search_result(self, path_list):
        # retrieve the directory path and filename respectively from the same input text fields text1 and text2
        directory = self.text1.text()
        filename = self.text2.text()
        # calculate the number of files found in the search operation
        file_number=len(path_list)
        if self.result_flag == 0:
            self.result = QLabel()
            # self.result.setStyleSheet('border: 1px solid #666;line-height:3em')
        # calculate the duration of the search operation
        duration = time.time() - self.search_begin_time
        # File not found in the search operation
        if not path_list:
            self.result.setText(f'<h2 style = "font-family: Arial">No file were founded In directory "{directory}" have name included with "{filename}"<br> In: {duration:.5}s')
        # find files in the search operation
        else:
            self.result.setText(f'<h2 style = "font-family: Arial">In directory "{directory}" which file have name included with "{filename}"<br>Founded: {file_number} file(s) in: {duration:.5}s</h2>' +'<font style="font-size:15px;font-family: Arial;">'+ '<br>'.join(path_list) +'</font>')
        if self.result_flag == 0:
            self.v_layout.addWidget(self.result)
            self.result_flag = 1
        # Delete previous search results before displaying new results
        if self.key_word_result_flag == 1:
            self.v_layout.takeAt(self.v_layout.count() - 1).widget().setParent(None)
            self.key_word_result_flag = 0
    # search for keywords in a given file or directory
    def keyword_find_begin(self):
        self.search_begin()
        if self.key_word_result_flag == 0:
            self.key_word_result = QLabel()
            # self.key_word_result.setStyleSheet('border: 1px solid #666;line-height:3em')
        self.keywordfind_begin_time = time.time()
        # filename = self.text2.text()  # detect through the input format of the second line
        filename = self.allpathes[0] if self.allpathes else self.text2.text()  # Replace the direct specification of input format in the second line with the format of the first element of self.allpaths
        typename = filename.split('.')[-1]
        keyword = self.text3.text()
        key_word = self.text3.text()
        if not key_word:
            QMessageBox.warning(self, "Warning", "Please enter a Keyword to find")
            return
        else:
            # The auxiliary functions (find_word(), find_excel(), find_text(), or find_pdf()) are used to search for specified keywords within files. They are utilized for searching keywords from specific types of files.
            if os.path.exists(self.text1.text()):
                if typename in ['.doc', '.docx','doc', 'docx']:
                    self.find_word(self.allpathes, keyword)
                elif typename in ['.xls', '.xlsx', 'xls', 'xlsx']:
                    self.find_excel(self.allpathes, keyword)
                elif typename in ['.txt', '.py', 'txt', 'py']:
                    self.find_text(self.allpathes, keyword)
                elif typename in ['.pdf', 'pdf']:
                    self.find_pdf(self.allpathes, keyword)
                # The warning message is displayed. The keyword is not specified or the specified file or directory does not exist.
                else:
                    QMessageBox.warning(self, 'Wrong file type!', 'Wrong file type, please input .docx/.pdf/.pdf/.py/.xlsx')
            else:
                QMessageBox.warning(self, 'File not exist ', 'File not exist, please input correct File name')
    # Display the results of keyword search in the widget
    def show_find_result(self, txt_list,Keyword,  duration):
        # Concatenate the matching file list into one string
        if txt_list:
            self.key_word_result.setText(f'<h2 style = "font-family: Arial">Founded keyword "{Keyword}"from above directory(s) in: {duration:.5}s</h2>' + '<font style="font-size:15px;font-family: Arial;">' + '<br/>'.join(self.txt_list) + '</font>')
        # including the number of matching keywords and the list of matching keyword directories.
        else:
            self.key_word_result.setText(f'<h2 style = "font-family: Arial">No word "{Keyword}" were founded from above directory in: {duration:.5}s</h2>')
        if self.key_word_result_flag == 0:
            self.v_layout.addWidget(self.key_word_result)
            self.key_word_result_flag = 1

    # Search for keywords from .docx file type
    def find_word(self, file_list, keyword):
        self.txt_list = []
        for word in file_list:
            if word.endswith('docx'):
                text = docx2txt.process(word)
                count = text.count(keyword)
                if count:
                    self.txt_list.append(f'Founded: {count:<5}  keyword(s) in ' + word)
        duration = time.time() - self.keywordfind_begin_time
        # Use show_find_result() to display the search results.
        self.show_find_result(self.txt_list,keyword,  duration)
    # Search for keywords in .xlsx file type.
    def find_excel(self, file_list, keyword):
        self.txt_list = []
        for excel in file_list:
            if excel.endswith('xls') or excel.endswith('xlsx'):
                try:
                    df = pd.read_excel(excel)
                    text = df.to_string()
                    count = text.count(keyword)
                    if count:
                        self.txt_list.append(f'<br>Founded: {count:<5}  keyword(s) in ' + excel)
                except Exception as e:
                    print(f"Error reading {excel}: {e}")
        duration = time.time() - self.keywordfind_begin_time
        # Use show_find_result() to display the search results.
        self.show_find_result(self.txt_list,keyword,  duration)
    # Search for keywords in .txt and .py file types.
    def find_text(self, file_list, keyword):
        self.txt_list = []
        for txt in file_list:
            if txt.endswith('txt') or txt.endswith('py'):
                with open(txt, 'r') as f:
                    text = f.read()
                count = text.count(keyword)
                if count:
                    self.txt_list.append(f'<br>Founded: {count:<5}  keyword(s) in ' + txt)

        duration = time.time() - self.keywordfind_begin_time
        # Use show_find_result() to display the search results.
        self.show_find_result(self.txt_list, keyword, duration)

    # Search for keywords from .pdf file type.
    def find_pdf(self, file_list, keyword):
        self.txt_list = []
        for pdf in file_list:
            if pdf.endswith('pdf'):
                with open(pdf, 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    pageNum = len(pdf_reader.pages)
                    text = '\n'.join([pdf_reader.pages[i].extract_text() for i in range(pageNum)])
                count = text.count(keyword)
                if count:
                    self.txt_list.append(f'<br>Founded: {count:<5}  keyword(s) in ' + pdf)
        duration = time.time() - self.keywordfind_begin_time
        # Use show_find_result() to display the search results
        self.show_find_result(self.txt_list, keyword, duration)

# The main function calls the QApplication() function for the required GUI page controls and instantiates the custom class File_search.
if __name__ == '__main__':
    app = QApplication()
    window = File_search()
    app.exec_()
