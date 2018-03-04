#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Feb 20 10:34:00 2018

@author: crantu
"""

import tkinter.filedialog as tf

import openpyxl as px
import zenhan as zh

import pandas as pd

import datetime
import re
import sys
import time
import traceback


class hp2sheet:
    def __init__(self):
        self.dt = datetime.datetime.today()
        # self.cell = "{c}{i}" # -> replaced with function
        # Character,Integer
        self.gakki = ""

    def main(self):
        try:
            f = open("./univName.txt", "r")
            self.univname = f.readline().strip()
            f.close()
        except:
            f.close()
            print("Error: univURL.txt")
            return 0

        try:    
            print("Please select your file.")
            print("\n")

            msg = "Are you sure to change this file?:\n{}"
            if self.openExcel() and \
                    self.yesno(msg.format(self.filename), True):
                pass  # Ask sure or not to change the file
            else:
                return 0
            print("\n")

            print("The downloaded html table will be converted to simple one.")
            print("If you do NOT want, please press ENTER.")

            if self.copy_input_to_temp(input("Which quoter?(1/2/3/4):")):
                input('\n"temp" will be converted.\n' +
                      "Please check and close your file then press ENTER.:")
            else:
                return 0

            msg = "\nDo you want to download classrooms?"
            if self.copy_tmp_to_forpdf() and self.yesno(msg, False):
                self.importclassroom()
                # change the file and ask want to download or not
                print("END!!")
            else:
                return 0
        except:
            return 0

        finally:
            return 0

    def openExcel(self):
        try:
            # root=tkinter.Tk()
            # root.withdraw()

            fTyp = [('Sheet copied from web site', '*.xlsx')]
            iDir = ".//"

            filename = tf.askopenfilename(filetypes=fTyp, initialdir=iDir)
            hantei = False

            if filename == "":
                print("The cancel butten was pushed.")
            else:
                wb = px.load_workbook(filename)
                wb.save(filename)

                self.filename = filename

                hantei = True
        except PermissionError:
            print("The file was not closed.")
        except:
            print("Error")
        finally:
            return hantei

    def copy_tmp_to_forpdf(self):
        try:
            wb = px.load_workbook(self.filename)

            ws1 = wb.get_sheet_by_name('ForPDF')
            ws2 = wb.get_sheet_by_name('temp')

            kogi_bango = ""
            # cell="{c}{i}"
            # Character,Integer

            for i in range(66, 71):
                for j1 in range(2, 17, 2):  # Kougi-bango #write course number
                    r1 = round(3*j1/2+2)
                    kogi_bango = ws2[self.cell(c=chr(i), i=j1)].value
                    if kogi_bango is None:
                        pass
                    elif type(kogi_bango) == str:
                        kogi_bango0 = kogi_bango.split("　", 1)[0].strip()
                        ws1[self.cell(c=chr(i), i=r1)].value = ""
                        ws1[self.cell(c=chr(i), i=r1)].value =\
                            zh.z2h(text=kogi_bango0, mode=7)
                    elif type(kogi_bango) == int:
                        kogi_bango0 = "{:06d}".format(kogi_bango)
                        ws1[self.cell(c=chr(i), i=r1)].value = ""
                        ws1[self.cell(c=chr(i), i=r1)].value =\
                            zh.z2h(text=kogi_bango0, mode=7)

                for j2 in range(3, 18, 2):  # write course title
                    r2 = round(((3*j2+7)/2)-1)
                    kogi_me = ws2[self.cell(c=chr(i), i=j2)].value
                    if kogi_me is not None:
                        kogi_me = kogi_me.replace("－", "-").\
                            replace("英語コミュニケーション", "EC")
                        ws1[self.cell(c=chr(i), i=r2)].value = ""
                        ws1[self.cell(c=chr(i), i=r2)].value =\
                            zh.z2h(text=kogi_me, mode=3)

            ws1["F2"].value = self.dt.strftime("%Y/%m/%d")
            if self.gakki != "":
                ws1["D1"].value = "Q{}".format(self.gakki)
            wb.save(self.filename)
            print("Successfully completed")
            return True

        except PermissionError:
            print("The file was not closed.")
            return False

        except:
            # print("Error")
            print("Unexpected error:", sys.exc_info())
            return False

    def importclassroom(self):
        try:
            wb = px.load_workbook(self.filename)
            ws1 = wb.get_sheet_by_name('ForPDF')
            # self.cell(c=chr(i),i=r1)

            cnobox = [[None for i0 in range(8)] for j0 in range(5)]
            for i1 in range(5):  # read data from the file
                for j1 in range(8):
                    cnobox[i1][j1] =\
                        ws1[self.cell(c=chr(i1+66), i=3*(j1+2)-1)].value

            cnolist = []
            for i in range(5):  # table->list
                for j in range(8):
                    inlist = False
                    for k in cnolist:
                        if cnobox[i][j] == k \
                          or cnobox[i][j] is None or cnobox[i][j] == "-":
                            inlist = True
                            break
                    if len(cnolist) == 0 and \
                            (cnobox[i][j] is None or cnobox[i][j] == "-"):
                        pass
                    elif inlist is False:
                        cnolist.append(cnobox[i][j])

            croomlist = []
            iii = len(cnolist)
            ii = 1
            print("Downloading...")
            for i2 in cnolist:  # import html
                print(ii, "of", iii)
                croomlist.append(self.findclassroom(i2))
                time.sleep(2)
                ii = ii+1

            for i in range(5):  # list->box
                for j in range(8):
                    inlist = False
                    m = 0
                    for k in cnolist:
                        if cnobox[i][j] is not None and cnobox[i][j] == k:
                            cnobox[i][j] = croomlist[m]
                            break
                        m = m+1

            for i1 in range(5):  # write data to the file
                for j1 in range(8):
                    if cnobox[i1][j1] is not None:
                        ws1[self.cell(
                            c=chr(i1+66), i=3*(j1+2))].value = ""
                        ws1[self.cell(
                            c=chr(i1+66), i=3*(j1+2))].value = cnobox[i1][j1]

            wb.save(self.filename)
            print("Successfully completed")
            return True
        except PermissionError:
            print("Please close the file.")
            return False
        except:
            # print("Error")
            # print("Unexpected error:", sys.exc_info())
            print("Unexpected error:", end="")
            traceback.print_exc()
            return False

    # download html and return classroom with course No
    def findclassroom(self, shozoku_jikanwari):
        try:
            nendo = self.dt.year
            if self.dt.month < 4:
                nendo = nendo-1
        #    nendo=2017
            if re.match(r"'.*", shozoku_jikanwari) is not None:
                shozoku_jikanwari = shozoku_jikanwari.replace("'", "")

            shozoku = shozoku_jikanwari[0:2]
            jikanwari = shozoku_jikanwari[2:6]

            obj_url = ("https://gs.{uname}.ac.jp/campusweb/campussquare.do?_"
                       "flowId=SYW4101101-flow&nendo={n}&"
                       "shozoku={s}&jikanwari={j}"
                       "&sylocale={lang}").format(uname=self.univname)
        #    obj_url="/slb_sample.html"

            try:
                df = pd.io.html.read_html(obj_url.format(
                    n=nendo, s=shozoku, j=jikanwari, lang="ja_JP"))
            except:
                df = pd.io.html.read_html(obj_url.format(
                    n=nendo, s=shozoku, j=jikanwari, lang="en_US"))

            return self.changeclassroom(df[0][1][6])
        except:
            return "error"

    def changeclassroom(self, classroom):
        if re.match(r".*,.*", classroom) is not None:
            cr = classroom
        elif classroom == "工学部１号館情報実習室１（CAE室）":
            cr = "工1-CAE室"
        elif re.match(r"一般教育棟.*", classroom) is not None:
            cr = classroom.replace("一般教育棟", "").replace("教室", "")
        elif re.match(r"工学部.*", classroom) is not None:
            cr = classroom.replace('工学部', "工").replace(
                "号館第", "-").replace("号館", "-").replace("講義室", "")
        elif re.match(r"情報実習室.*", classroom) is not None:
            cr = classroom.replace("情報実習室", "情")
        elif re.match(r"理学部.*", classroom) is not None:
            cr = classroom.replace("理学部", "理").replace(
                "号館第", "-").replace("号館", "-").replace("講義室", "")
        else:
            cr = classroom

        cr = cr.replace(" ", "")

        return zh.z2h(text=cr, mode=3)

    def copy_input_to_temp(self, gakki):
        try:
            wb = px.load_workbook(self.filename)

            ws2 = wb.get_sheet_by_name('temp')
            ws3 = wb.get_sheet_by_name('input')

            # string=""

            if gakki == "1" or gakki == "3":
                for i in range(5):
                    for j in range(8):  # Kougi-bango #write course number
                        ws2[self.cell(
                            c=chr(i+66), i=2*(j+1))].value = ""
                        ws2[self.cell(c=chr(i+66), i=2*(j+1))].value =\
                            ws3[self.cell(c=chr(i+67), i=8*j+1)].value
                        ws2[self.cell(
                            c=chr(i+66), i=2*(j+1)+1)].value = ""
                        ws2[self.cell(c=chr(i+66), i=2*(j+1)+1)].value =\
                            ws3[self.cell(c=chr(i+67), i=8*j+2)].value
                msg = "Successfully completed"
                self.gakki = gakki

            elif gakki == "2" or gakki == "4":
                for i in range(5):
                    for j in range(8):  # Kougi-bango #write course number
                        ws2[self.cell(
                            c=chr(i+66), i=2*(j+1))].value = ""
                        ws2[self.cell(c=chr(i+66), i=2*(j+1))].value =\
                            ws3[self.cell(c=chr(i+67), i=8*j+3)].value
                        ws2[self.cell(
                            c=chr(i+66), i=2*(j+1)+1)].value = ""
                        ws2[self.cell(c=chr(i+66), i=2*(j+1)+1)].value =\
                            ws3[self.cell(c=chr(i+67), i=8*j+4)].value
                msg = "Successfully completed"
                self.gakki = gakki

            else:
                msg = 'The sheet "temp" was not changed.'

            wb.save(self.filename)
            print(msg)
            return True

        except PermissionError:
            print("The file was not closed.")
            return False

        except:
            print("Error")
            return False

    def yesno(self, msg, y_n0=None):
        ans = None
        while ans is None:
            if y_n0 is True:
                ans = True
                y_n = "[y]"
            elif y_n0 is False:
                ans = False
                y_n = "[n]"
            elif y_n0 is None:
                y_n = ""
            else:
                raise

            ans0 = input(msg+"(y/n)"+y_n+":")
            if ans0 == "y":
                ans = True
            elif ans0 == "n":
                ans = False
            elif ans0 == "":
                pass
            else:
                ans = None
        return ans

    def cell(self, c, i):
        return "{c}{i}".format(c=c, i=i)


def quick_start():
    h2s = hp2sheet()
    h2s.main()


if __name__ == "__main__":
    quick_start()
