# -*- coding: utf-8 -*-
"""
Created on Fri Jul 21 08:58:34 2017

@author: okayama_univ
"""

import tkinter
from tkinter import filedialog as tkFileDialog

import openpyxl as px

class conv4jpg:
    def __init__(self):
        pass
    
    def main(self):
        try:
            print("Please select your file.")
            print("\n")
            
            if self.main0()==True and \
            self.yesno("Are you sure to change this file?:\n"+self.filename,True)==True: 
                self.main4()#Ask sure or not to change the file
            else:
                return 0
        except:
            return 0        
        finally:
            return 0

    def main0(self):
        try:
            root=tkinter.Tk()
            root.withdraw()
            
            fTyp=[('Sheet copied that you wnat to convert','*.xlsx')]    
            iDir=".//"
    
            filename=tkFileDialog.askopenfilename(filetypes=fTyp,initialdir=iDir)
            hantei=False
            
            if filename=="":
                print("The cancel butten was pushed.")
            else:
                wb = px.load_workbook(filename)
                wb.save(filename)
    
                self.filename=filename  #self.filename="F:\Okadai\時間割\時間割1Q2_1S\時間割_170530.xlsx"#test!!
                hantei=True              
        except PermissionError:
            print("The file was not closed.")
        except:
            print("Error")
        finally:
            return hantei
        
    def main4(self):
        try:
            wb = px.load_workbook(self.filename)
           
            ws1 = wb.get_sheet_by_name('ForPDF')
            ws1_2 = wb.get_sheet_by_name('ForJPG')

            for i in range(5):
                for j in range(24):
                    ws1_2[chr(i+67)+str(j+6)].value=ws1[chr(i+66)+str(j+5)].value
            
            ws1_2["G3"].value=ws1["F2"].value
            ws1_2["E2"].value=ws1["D1"].value
            
            wb.save(self.filename)
            print("Successfully completed")
            return True
        
        except PermissionError:
            print("The file was not closed.")
            return False
            
        except:
            print("Error")
            return False    
    
    def yesno(self,msg,y_n0=None):
        ans=None
        while ans==None:
            if y_n0==True:
                ans=True
                y_n="[y]"
            elif y_n0==False:
                ans=False
                y_n="[n]"
            elif y_n0==None:
                y_n=""
            else:
                raise
                    
            ans0=input(msg+"(y/n)"+y_n+":")           
            if ans0=="y":
                ans=True
            elif ans0=="n":
                ans=False
            elif ans0=="":
                pass
            else:
                ans=None        
        return ans
    
def quick_start():
    c4j=conv4jpg()
    c4j.main()

quick_start()