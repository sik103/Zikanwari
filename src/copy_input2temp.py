#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Apr  6 12:56:20 2018

@author: crantu
"""

import openpyxl as px
from .cell_for_openpyxl import cell


def copy_input_to_temp(filename, gakki):
    try:
        wb = px.load_workbook(filename)
        ws2 = wb['temp']
        ws3 = wb['input']

        if gakki == "1" or gakki == "3":
            dj = 0

        elif gakki == "2" or gakki == "4":
            dj = 2

        else:
            print("ValueError: Input 1, 2, 3, or 4")
            return False

        for i in range(5):
            for j in range(8):  # Kougi-bango #write course number
                ws2[cell(
                    c=chr(i+66), i=2*(j+1))].value = ""
                ws2[cell(c=chr(i+66), i=2*(j+1))].value =\
                    ws3[cell(c=chr(i+67), i=8*j+1+dj)].value
                ws2[cell(
                    c=chr(i+66), i=2*(j+1)+1)].value = ""
                ws2[cell(c=chr(i+66), i=2*(j+1)+1)].value =\
                    ws3[cell(c=chr(i+67), i=8*j+2+dj)].value
        msg = "Successfully completed"

        wb.save(filename)
        print(msg)
        return gakki

    except PermissionError:
        print("The file was not closed.")
        return False


if __name__ == "__main__":
    copy_input_to_temp("/home/crantu/Desktop/a.xlsx", "1")
