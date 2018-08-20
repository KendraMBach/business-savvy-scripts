import pandas as pd
import xlsxwriter
import xlrd
import os
import glob
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showerror
import sys
import warnings
warnings.filterwarnings("always")

# GUI window
window = Tk()
window.title("File Input")
window.geometry('600x200')

btn_text_wbpath = tk.StringVar()
btn_text_wbpath.set("Browse")

btn_text_filepath = tk.StringVar()
btn_text_filepath.set("Browse")

def browsedir():
    global file_path
    filedir = filedialog.askdirectory()
    file_path = str(filedir)
    btn_text_filepath.set(str(file_path))

def browsewbpath():
    global wb_path
    wbdir = filedialog.askdirectory()
    wb_path = str(wbdir)
    btn_text_wbpath.set(str(wb_path))


file_path_label = Label(window, text="Please provide the folder path of your Excel files:")
file_path_label.grid(column=0, row=0)

file_path_entry = Button(window,textvariable = btn_text_filepath, width=20, command=browsedir)
file_path_entry.grid(column=1, row=0)

new_workbook_path_label = Label(window, text="Where would you like to store the new workbook?")
new_workbook_path_label.grid(column=0, row=1)

new_workbook_path_entry = Button(window, textvariable = btn_text_wbpath, width=20, command=browsewbpath)
new_workbook_path_entry.grid(column=1, row=1)

new_workbook_name_label = Label(window, text="What would you like to name the new workbook?")
new_workbook_name_label.grid(column=0, row=2)

new_workbook_name_entry = Entry(window,width=20)
new_workbook_name_entry.grid(column=1, row=2)


def clicked():
    # Create a Pandas Excel writer using XlsxWriter as the engine
    wb_name = new_workbook_name_entry.get()
    writer = pd.ExcelWriter((wb_path + "/" + wb_name + ".xlsx"), engine='xlsxwriter')

    # Attach each existing Excel document to the same workbook as separate worksheets
    # Names of the files are retained as worksheet names
    for path in glob.glob(file_path + "/*.xls"):
        ds_name = str(os.path.basename(path))
        ds = pd.read_excel(path)
        ds.to_excel(writer, sheet_name=ds_name, index = False)

    #Save and close writer to allow changes to new workbook to persist
    writer.save()
    writer.close()
    #Close window

btn = Button(window, text="Submit", command=clicked)
btn.grid(column=0, row=3)
window.mainloop()
