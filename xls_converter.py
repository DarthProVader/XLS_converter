import os
import win32com.client as win32
from tkinter import messagebox

#find .xls file in current working directory
file = []
for i in os.listdir():
    if i.endswith(".xls"):
        file.append(i)

fname = os.getcwd() +"\\" + file[0] #create path to this file
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)

wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()

messagebox.showinfo("Message:", "Done!")
