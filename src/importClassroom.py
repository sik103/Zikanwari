import openpyxl as px
from .cell_for_openpyxl import cell
from .getClassroom import GetClassroom
import time
from datetime import datetime

import os
from os import path
os.chdir(path.dirname(path.abspath(__file__)))


def importclassroom(filename, debug_mode=False):
    dt = datetime.today()
    gc = GetClassroom(dt)
    try:
        wb = px.load_workbook(filename)
        ws1 = wb['ForPDF']

        cnobox = [[None for i0 in range(8)] for j0 in range(5)]
        for i1 in range(5):  # read data from the file
            for j1 in range(8):
                cnobox[i1][j1] =\
                    ws1[cell(c=chr(i1 + 66), i=3 * (j1 + 2) - 1)].value

        cnolist = []
        for i in range(5):  # table->list
            for j in range(8):
                inlist = False
                for k in cnolist:
                    if cnobox[i][j] == k \
                            or not cnobox[i][j] or cnobox[i][j] == "-":
                        inlist = True
                        break
                if len(cnolist) == 0 and \
                        (not cnobox[i][j] or cnobox[i][j] == "-"):
                    pass
                elif inlist is False:
                    cnolist.append(cnobox[i][j])

        croomlist = []
        iii = len(cnolist)
        print("Downloading...")
        for ii, cnol in enumerate(cnolist):  # import html
            print("{} of {}".format(ii + 1, iii))
            croomlist.append(gc.getClassroom(cnol, debug_mode))
            time.sleep(5)

        for i in range(5):  # list->box
            for j in range(8):
                for m, cnol in enumerate(cnolist):
                    if cnobox[i][j] and cnobox[i][j] == cnol:
                        cnobox[i][j] = croomlist[m]
                        break

        for i1 in range(5):  # write data to the file
            for j1 in range(8):
                if cnobox[i1][j1]:
                    ws1[cell(c=chr(i1 + 66), i=3 * (j1 + 2))].value = ""
                    ws1[cell(c=chr(i1 + 66), i=3 * (j1 + 2))].value =\
                        cnobox[i1][j1]

        wb.save(filename)
        print("Successfully completed")
        return True

    except PermissionError:
        print("Please close the file.")
        return False


if __name__ == "__main__":
    importclassroom("../files_for_debug/a.xlsx", True)
