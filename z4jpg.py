# -*- coding: utf-8 -*-
"""
Created on Fri Jul 21 08:58:34 2017

@author: okayama_univ
"""

# import tkinter
# from tkinter import filedialog as tkFileDialog
import openpyxl as px
from src.yesno_interface import yesno
from src.check_excelfile import CheckExcelFile

cef = CheckExcelFile()


class conv4jpg:
    def __init__(self):
        pass

    def main(self):
        try:
            print("Please select your file.")
            print("\n")

            if cef.openExcel():
                self.filename = cef.filename
                if yesno("Are you sure to change this file?:\n"+self.filename,
                         True):
                    self.main4()  # Ask sure or not to change the file
            else:
                return 0
        except PermissionError:
            return 0
        finally:
            return 0

    def main4(self):
        try:
            wb = px.load_workbook(self.filename)

            ws4pdf = wb.get_sheet_by_name('ForPDF')
            ws4jpg = wb.get_sheet_by_name('ForJPG')

            for i in range(5):
                for j in range(24):
                    ws4jpg[chr(i+67)+str(j+6)].value\
                        = ws4pdf[chr(i+66)+str(j+5)].value

            ws4jpg["G3"].value = ws4pdf["F2"].value
            ws4jpg["E2"].value = ws4pdf["D1"].value

            wb.save(self.filename)
            print("Successfully completed")
            return True

        except PermissionError:
            print("The file was not closed.")
            return False


if __name__ == "__main__":
    c4j = conv4jpg()
    c4j.main()
