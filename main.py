import os, sys
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
import threading
import datetime

from docx import Document # Wordprocess Reader 전용


UI_add = "./UI/Doc_Finder.ui"

class MainDialog(QMainWindow):
    def __init__(self):
        QDialog.__init__(self, None)
        uic.loadUi(UI_add,self)
        self.PathSelectBtn.clicked.connect(self.PathSelectEvemt)
        self.Search_Btn.clicked.connect(self.SearchStart)
        self.setFixedSize(450, 330)
        self.fname = None # directory Path
        self.load_data()
        self.total_detectresult = []
        
    def load_data(self):
        try:
            f = open("./temp_data.dat",'r+')
            line = f.readline()
            if not line: return f.close()
            self.Path_label.setText(line)
            f.close()
        except:
            print("Can't find 'temp_data'")

    def save_data(self):
        f = open("./temp_data.dat",'w+')
        f.write(self.Path_label.text())
        f.close()

    def PathSelectEvemt(self):
        self.fname = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.Path_label.setText(self.fname)
        

    def split_finder(self,find_target_list):
        checked_list = ['docx','pdf','xlsx']
        detect_result=[]
        detected_count = 0
        for target in find_target_list:
                for extneder in checked_list: # 선택된 확장자 내에서 검색
                    if((target.split('.')[-1]).lower() == extneder):
                        
                        if(extneder == 'docx'):
                            try:
                                document = Document(target)
                                for par in document.paragraphs:  # to extract the whole text
                                    if(par.text.find(self.target_textbox.text()) != -1):
                                        detect_result.append(target)
                                        break
                            except:
                                break
                        elif(extneder == 'pdf'): 
                            '''
                            # open the pdf file
                            object = PyPDF2.PdfFileReader("test.pdf")

                            # get number of pages
                            NumPages = object.getNumPages()

                            # define keyterms
                            String = "Social"

                            # extract text and do the search
                            for i in range(0, NumPages):
                                PageObj = object.getPage(i)
                                print("this is page " + str(i)) 
                                Text = PageObj.extractText() 
                                # print(Text)
                                ResSearch = re.search(String, Text)
                                print(ResSearch)
                                
                            reader = PyPDF2.PdfFileReader('Complete_Works_Lovecraft.pdf')
                            reader.getPage(7-1).extractText()
                            '''
                        elif(extneder == 'xlsx'): 
                            pass
                        else:
                            pass
                        
                        try:
                            detected_count +=1
                            result = open(target, 'r', encoding='EUC-KR').read().find(self.target_textbox.text()) # 'MOVE_PICK'
                            if(result != -1):
                                detect_result.append(target)
                        except:
                            pass
        self.total_detectresult.append(detect_result)
        
    def SearchStart(self):
        self.total_directory = []
        for (path, dir, files) in os.walk(self.Path_label.text()):
            for i in range(len(files)):
                self.total_directory.append(path+'/'+files[i])
        s_t = datetime.datetime.now()
        Thred1 = threading.Thread(target= self.split_finder, args =([self.total_directory[:len(self.total_directory)//2]]))
        Thred2 = threading.Thread(target= self.split_finder, args =([self.total_directory[len(self.total_directory)//2:]]))
        Thred1.start()
        Thred2.start()
        Thred1.join()
        Thred2.join()
        self.total_detectresult = sum(self.total_detectresult,[])
        e_t = datetime.datetime.now()
        self.total_count.setText("파일개수 : "+str(len(self.total_detectresult)))
        self.find_timer.setText("탐색시간 : "+str(e_t-s_t))
        '''
        Drawing Table Start
        '''
        self.Result_table.setRowCount(0)
        if(len(self.total_detectresult) > 0):
            # Finish Detect -> Make Table
            self.Result_table.setRowCount(len(self.total_detectresult))
            self.Result_table.setColumnCount(2)
            self.Result_table.setHorizontalHeaderLabels(["file_name", "Path"])
            for i in range(len(self.total_detectresult)):
                file_name = self.total_detectresult[i].split('/')[-1]
                self.Result_table.setItem(i,0, QTableWidgetItem(file_name))
                self.Result_table.setItem(i,1, QTableWidgetItem(self.total_detectresult[i][:-1*len(file_name)]))
            # table widget Setting
            header = self.Result_table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.Stretch)
            self.save_data()
        else:
            QMessageBox.information(self,"Information", "Can not find anything.")
        self.total_detectresult = [] # Buffer Initialize for Next Search
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    Main_Dialog = MainDialog()
    Main_Dialog.show()
    app.exec_()