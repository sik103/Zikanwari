#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Apr  6 12:43:13 2018

@author: crantu
"""

import openpyxl as px
import zenhan as zh
from cell_for_openpyxl import cell
from datetime import datetime

import os
from os import path
os.chdir(path.dirname(path.abspath(__file__)))


def copy_tmp_to_forpdf(filename, gakki):
    dt = datetime.today()
    try:
        wb = px.load_workbook(filename)

        ws1 = wb['ForPDF']
        ws2 = wb['temp']

        kogi_bango = ""

        for i in range(66, 71):
            for j1 in range(2, 17, 2):  # Kougi-bango #write course number
                r1 = round(3*j1/2+2)
                kogi_bango = ws2[cell(c=chr(i), i=j1)].value
                if kogi_bango is None:
                    pass
                elif type(kogi_bango) == str:
                    kogi_bango0 = kogi_bango.split("　", 1)[0].strip()
                    ws1[cell(c=chr(i), i=r1)].value = ""
                    ws1[cell(c=chr(i), i=r1)].value =\
                        zh.z2h(text=kogi_bango0, mode=7)
                elif type(kogi_bango) == int:
                    kogi_bango0 = "{:06d}".format(kogi_bango)
                    ws1[cell(c=chr(i), i=r1)].value = ""
                    ws1[cell(c=chr(i), i=r1)].value =\
                        zh.z2h(text=kogi_bango0, mode=7)

            for j2 in range(3, 18, 2):  # write course title
                r2 = round(((3*j2+7)/2)-1)
                kogi_me = ws2[cell(c=chr(i), i=j2)].value
                if kogi_me is not None:
                    kogi_me = kogi_me.replace("－", "-").\
                        replace("英語コミュニケーション", "EC")
                    ws1[cell(c=chr(i), i=r2)].value = ""
                    ws1[cell(c=chr(i), i=r2)].value =\
                        zh.z2h(text=kogi_me, mode=3)

        ws1["F2"].value = dt.strftime("%Y/%m/%d")
        if gakki != "":
            ws1["D1"].value = "Q{}".format(gakki)
        wb.save(filename)
        print("Successfully completed")
        return True

    except PermissionError:
        print("The file was not closed.")
        return False


if __name__ == "__main__":
    copy_tmp_to_forpdf("/home/crantu/Desktop/a.xlsx", "1")
