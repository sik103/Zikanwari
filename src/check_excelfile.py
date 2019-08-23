#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun 15 13:10:45 2018

@author: crantu
"""

import openpyxl as px
import sys
from pathlib import Path

import os
from os import path
os.chdir(path.dirname(path.abspath(__file__)))


class CheckExcelFile:
    def __init__(self):
        pass

    def openExcel(self):
        if len(sys.argv) == 2:
            filename = sys.argv[1]
        else:
            import tkinter
            import tkinter.filedialog as tf
            root=tkinter.Tk()
            root.withdraw()

            fTyp = [('Okadai Timetable', '*.xlsx')]
            iDir = path.join(str(Path.home()), "Desktop")

            filename = tf.askopenfilename(filetypes=fTyp, initialdir=iDir)
        try:
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
