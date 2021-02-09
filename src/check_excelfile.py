import openpyxl as px
import tkinter.filedialog as tf

import os
from os import path
os.chdir(path.dirname(path.abspath(__file__)))


class CheckExcelFile:
    def __init__(self):
        pass

    def openExcel(self):
        try:
            # root=tkinter.Tk()
            # root.withdraw()

            fTyp = [('Okadai Timetable', '*.xlsx')]
            iDir = ".//"

            filename = tf.askopenfilename(filetypes=fTyp, initialdir=iDir)
            # hantei = False

            if filename == "":
                print("The cancel butten was pushed.")
            else:
                wb = px.load_workbook(filename)
                wb['ForPDF']
                wb['ForJPG']
                wb['temp']
                wb['input']
                wb.save(filename)

                self.filename = filename

                return True
        except PermissionError:
            print("The file was not closed.")
            return False
        except KeyError:
            print("The Worksheets were wrong. ")
            return False


if __name__ == "__main__":
    cef = CheckExcelFile()
    cef.openExcel()
