#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Feb 20 10:34:00 2018

@author: crantu
"""

from copy_tmp2forpdf import copy_tmp_to_forpdf
from copy_input2temp import copy_input_to_temp
from yesno_interface import yesno
from importClassroom import importclassroom
from check_excelfile import CheckExcelFile

cef = CheckExcelFile()


class hp2sheet:
    def __init__(self):
        pass

    def main(self):
        print("Please select your file.")
        print("\n")

        msg = "Are you sure to change this file?:\n{}"
        if cef.openExcel():
            self.filename = cef.filename
            if not yesno(msg.format(self.filename), True):
                return 0  # Ask sure or not to change the file
        else:
            return 0
        print("\n")

        print("The downloaded html table will be converted to simple one.")
        print("If you do NOT want, please press ENTER.")

        gakki = copy_input_to_temp(self.filename,
                                   input("Which quoter?(1/2/3/4):"))
        if gakki:
            input('\n"temp" will be converted.\n' +
                  "Please check and close your file then press ENTER.:")
        else:
            return 0

        msg = "\nDo you want to download classrooms?"
        if copy_tmp_to_forpdf(self.filename, gakki) and yesno(msg, False):
            importclassroom(self.filename)
            # change the file and ask want to download or not
            print("END!!")
        else:
            return 0


if __name__ == "__main__":
    h2s = hp2sheet()
    h2s.main()
