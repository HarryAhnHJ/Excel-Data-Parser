import pathlib
import shutil
import openpyxl as xl
import os
import vars
import fee
import pandas as pd
import traceback
import sys, os


#print(traceback.format_exc())

def transformFile(file: str,qtr: str,year: str)->bool:
    '''
    Fucntion for individual excel files in the folder. 
    Get new file path & file name
    Replace new path & name with existing path & name
    '''

    newfilename = getnewfilename(file,qtr,year)

    if newfilename != "":
        newfilepath = getnewfilepath(file,qtr,year)
        #changing file path, only if new name is viable
        newfile = os.path.join(newfilepath,newfilename).replace("\\","/")
        shutil.move(file, newfile)
        return True
    else:
        return False

def getnewfilename(file: str,qtr: str,year: str)->str:
    '''
    create new file name with fetched venture name and yr-qtr
    '''
    venture_info = fetch_venture_name(file,qtr,year)

    if len(venture_info) == 0:
        print("Error with the fee worksheet.")
        return ""
    elif venture_info[0] is None:
        return ""
    
    prefix = venture_info[0]
    suffix = " - Q" + str(qtr) + venture_info[1]
    
    filename = prefix + suffix
    return filename


def getnewfilepath(file: str,qtr: str,year: str)->str:
    '''
    get file path of new destination
    '''
    # target_dir = (
    #     str(pathlib.Path.home()) + "/OneDrive - Quadreal Property Group" + od_path
    # )
    target_dir = (
        str(pathlib.Path.home()) + "/OneDrive - Quadreal Property Group" + vars.test_path
    ) # test path, comment this out and uncomment above for real application

    newFolderName = year + " Q" + qtr
    # print("quarter is " + qtr)
    path = os.path.join(target_dir,newFolderName).replace("\\","/")
    
    if not os.path.exists(path):    
        os.mkdir(path)
    else:
        print("Target directory already exists.")
    return path

#FUNCTION BELOW & FEE FUNCTION NEEDS TO BE UPDATED:
#1. MULTIPLE DUPLICATES CREATED IN OUTPUT - SHOULD BE ONLY ONE
#2. AS OF NOW, FUNCTION ONLY WORKS IF CAPITAL DEPLOYMENT SHEET IS FIRST - ORDER SHOULD NOT MATTER

def fetch_venture_name(file: str,qtr: str,year: str)->list[str]:
    '''
    Get venture name and type of report (AM Fee or Deployment forecast or both) 
        - If neither report exists, return empty
        - If only AM Fee and venture name exists, record fee and create new file name
        - If only Deployment Forecast and venture name exists, create new file name
        - If either/both exist but no name or wrong name, flag error and provide new name in UI
    '''
    try:
        wb = xl.load_workbook(filename = file,data_only=True)   
    except:
        print("Probably not an excel file. Ignoring this file...")
        return []

    QR_venture_name = ""
    deployment = False # this means deployment sheet exists, but not necesarily the correct venture name
    am_fee = False # this means AM fee sheet exists, but not necessarily with the correct venture name
    am_ws = ""

    '''
    Checks each worksheet to see if capital deployment sheet or AM fee sheet exist
    '''
    for ws in wb:
        if ws.sheet_state == "visible":
            venture_name = ""
            if str(ws.cell(row=4,column=1).value) == "INVESTMENT NAME:":
                # print("capital deployment sheet is visible!!")
                if str(ws.cell(row=4,column=2).value) != "":
                    venture_name = str(ws.cell(row=4,column=2).value).strip()
                deployment = True
            elif str(ws.cell(row=6,column=1).value) == "INVESTMENT NAME:":
                if str(ws.cell(row=6,column=2).value) != "":
                    venture_name = str(ws.cell(row=6,column=2).value).strip()
                am_fee = True
                am_ws = ws.title
        if venture_name != "":
            break
        
    '''
    If no relevant worksheet is found, leave the file alone & flag
    If worksheet is found, get the name of the venture using the cell found above
    '''
    qr_name_temp = ""
    if not am_fee and not deployment:
        print("Expected worksheet not found. Check if they submitted with the correct template.")
        return []
    else:
        qr_name_temp = vars.venture_names.get(venture_name.lower())
    
    '''
    
    '''
    if am_fee:
        if (qr_name_temp != "") & (qr_name_temp is not None):
            QR_venture_name = str(qr_name_temp).strip()
            print("QR venture name is " + str(QR_venture_name))
            try:
                fee.recordfee(QR_venture_name,file,am_ws,qtr,year)
            except Exception as e:
                print(traceback.format_exc())
                print(e)
                return []
        else:
            print("Venture name not found. Wrong name or cell is empty")

    print("No more worksheets to look through")
    
    filename = ""
    if deployment & am_fee:
        filename = " QRI Capital Deployment Forecast and AM Fee.xlsx"
    elif deployment:
        filename = " QRI Capital Deployment Forecast.xlsx"
    elif am_fee:
        filename = " QRI AM Fee.xlsx"
    elif QR_venture_name is not None:
        filename = ""
    else:
        return []
    return [QR_venture_name,filename]


def exception_ventures():
    print("Venture Exception handling not implemented yet, exiting now...")