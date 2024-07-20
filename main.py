
from tkinter import ttk,StringVar
from tkinter import filedialog
from tkinterdnd2 import TkinterDnD
from tkinter import Toplevel
from datetime import date
import os

import openpyxl as xl
import os

import vars
import fee
import transform

def browseFiles():
    '''
    search for folder where AM Fee files are locally saved, 
    then call rename and move file methods for each file in folder
    '''
    source_dir = filedialog.askdirectory(
        initialdir = vars.initfile,
        title = "Select a Folder!")

    if not source_dir:
        print("File not found. Please try again!")
    
    year = getYear()
    qtr = getQuarter()

    files = os.listdir(source_dir)
    print(files)

    #number of files needed to be mapped & recorded
    num_files = len(files)
    #counting number of files that have been mapped & recorded
    cnt_files = 0

    for file in files:
        file = source_dir + "/" + file
        print("            Currently working on: " + file)

        if file.endswith(('.xlsx','.csv','xlsm','xls')) and not os.path.isdir(file):
            transform_fix = transform.transformFile(file,qtr,year) 
            if transform_fix != []: #should only return non-empty list if ther needs to be venture name fix
                newfilepath = transform_fix[0]
                prefix = transform_fix[1]
                suffix = transform_fix[2]
                print("venture name error:")
                name_error(prefix)
        else:
            print("File done")
        cnt_files += 1 

    print("Renamed and moved all files. Exporting master fee dataframe...")

    if cnt_files == num_files:
        print("Last fee added. Exporting completed fee table:")
        fee.export_fee_db()
    else:
        print("Error: Not all files have been parsed. Check output and try again.")

    exit()
    

def name_error(prefix: str):#!!!
    '''
    error ui when it cannot determine name of venture from fee tab
    '''
    errorWindow = Toplevel(root)
    errorWindow.title("Unknown Venture Name")
    errorWindow.geometry("300x300")

    sub_frm = ttk.Frame(errorWindow, padding=10)

    new_venture_name = StringVar()
    
    init_label    = ttk.Label(
    sub_frm,
    text="Please correct the following venture name found in the Partner's report:",
    anchor="center").grid(columnspan=4, row=1)

    name_label    = ttk.Label(
    sub_frm,
    text=prefix).grid(column=1,row=2)

    input_qtr_text = ttk.Label(
    sub_frm,
    text = "Enter Correct QR Venture Name:"
    ).grid(column=1,row=3)

    input_qtr = ttk.Entry(
    sub_frm,
    textvariable=new_venture_name
    ).grid(column=2,row=3)

    input_confirm = ttk.Button(
    sub_frm,
    text="Submit",
    command=getnewname(new_venture_name)).grid(column=3,row=3)

    return

def getnewname(name: str):
    newname = new_venture_name.get()

def getQuarter()->str:
    '''
    set quarter based on tkinter input and return the value
    '''
    qtr = qtr_var.get() 

    try:
        if (int(qtr) < 1) | (int(qtr) > 4):
            print("Error: Invalid Quarter Input")
    except:
        print("Pleae input a valid number for Quarter")
    print("Date set as Q" + qtr)
    return qtr
    
    
def getYear()->str:
    '''
    set year based on current year and return the value
    '''
    today = date.today()
    # year = today.strftime("%Y")
    year = str(2024)
    return year
    

#-------------------------------------------------------------
'''
UI tkinter below
'''

root = TkinterDnD.Tk()
root.title("AM Fee Automation Tool")
# root.geometry('400x250')

frm = ttk.Frame(root, padding=10)
frm.grid()

qtr_var = StringVar()

init_label    = ttk.Label(
    frm,
    text="Select the corresponding Quarter you would like to input data for.\nThen, press 'Browse' to select the files you would like to parse through.",
    anchor="center").grid(columnspan=4, row=1)

blank_label    = ttk.Label(
    frm,
    text="").grid(column=1,row=2)

input_qtr_text = ttk.Label(
    frm,
    text = "Enter Quarter:"
    ).grid(column=1,row=3)

input_qtr = ttk.Entry(
    frm,
    textvariable=qtr_var
    ).grid(column=2,row=3)

input_confirm = ttk.Button(
    frm,
    text="Submit",
    command=getQuarter).grid(column=3,row=3)

browse_button = ttk.Button(
    frm,
    text="Browse",
    command=browseFiles).grid(column=2, row=4)

quit_button   = ttk.Button(
    frm,
    text="Quit",
    command=root.destroy).grid(column=2,row=5)
    
root.mainloop()
