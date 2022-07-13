# IMPORTS ############################################################
import sys
import camelot as cam

from shutil import copyfile

# imports for PDF_path() method
import datetime
from table_pdf import Table_to_pdf

# Imports for creating excel file (create_excel_file() method)
from win32com.client import Dispatch
import os


# Imports for user interface
from PyQt5.QtWidgets import (
    QApplication, QMainWindow,
)
from PyQt5.uic import loadUi
from interface import *


class Window(QMainWindow, Ui_MainWindow):
    # SHOW MAIN WINDOW ######################################################
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.setWindowTitle('Control de Calidad')


   # RED BUTTON TO GET PATH PDF ############################################
        self.pushButton.clicked.connect(self.PDF_path)

    #GREEN BUTTON TO CREATE EXCEL FILE #######################################
        self.pushButton_2.clicked.connect(self.create_excel_file)



    # GET PDF PATH ##########################################################
    def PDF_path(self):

        filename = QtWidgets.QFileDialog.getOpenFileName()
        path = filename[0]
        print(path)


        # SPLITTING THE PATH FILE AND GETTING JUST THE NAME FILE IN "pdf_path" VARIABLE #############

        list_path = path.split('/')
        pdf_path = list_path[-1]




        # EXTRACTING TABLE AND CONVERTING TO DATA FRAME
        table = Table_to_pdf(path)
        self.table = table.extract_table()

        # INSERTING EXCEL NAME TO THE lineEdit OBJECT  ##########################

        pdf_pathFile = pdf_path.replace('.pdf', '') # <--- removing ".pdf" from path name

        self.excel_name = f'{pdf_pathFile} - {datetime.datetime.now().strftime("%d %b %Y ")}'
        self.lineEdit.setText(self.excel_name)
 #
 #Creating a copy excel file  ###########################################

    def create_excel_file(self):

        ### sourceFile path

        # path1 = r'C:\Users\alex2\Downloads\DS Projects\ControlDeCalidad\Formato_Calidad.xlsx'
        path1 = os.path.abspath("Formato_Calidad.xlsx")
        print(path1)

        ### new excel file path

        desktop = os.path.expanduser("~/Desktop")


        # Reading the workbook  (path1)
        xl = Dispatch("Excel.Application")


        # wb1 = xl.Workbooks.Open(Filename=path1)



        # Coping the excel file
        copyfile(path1, f"{desktop}/{self.lineEdit.text()}.xlsx")

        # new excel path
        path2 = f"{desktop}/{self.lineEdit.text()}.xlsx"

        # Open the excel file
        wb2 = xl.Workbooks





        # sheet of spacifications of new excel file
        Ws5 = wb2.Worksheets(5)



        # self.table.insert(0,'Variables')

        # Pasting table extracted from pdf file to ws5
        Ws5.Range(Ws5.Cells(1, 1),  # Cell to start the "paste"
                 Ws5.Cells(1 + len(self.table.index) - 1,
                          1 + len(self.table.columns) )
                 ).Value = self.table.to_records()








        xl.Quit()

        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = Window()
    win.show()
    sys.exit(app.exec())
