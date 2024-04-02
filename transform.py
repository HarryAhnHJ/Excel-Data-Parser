import pathlib
import shutil
import openpyxl as xl
import os
import vars
import fee
import pandas as pd
import traceback


#print(traceback.format_exc())

def transformFile(file: str,qtr: str,year: str):
    '''
    Fucntion for individual excel files in the folder. 
    Get new file path & file name
    Replace new path & name with existing path & name
    '''

    newfilename = getnewfilename(file,qtr,year)
    newfilepath = getnewfilepath(file,qtr,year)



    #changing file path
    newfile = os.path.join(newfilepath,newfilename).replace("\\","/")
    shutil.move(file, newfile)


def getnewfilename(file: str,qtr: str,year: str)->str:
    '''
    create new file name with fetched venture name and yr-qtr
    '''
    venture_info = fetch_venture_name(file,qtr,year)

    if len(venture_info) == 0:
        print("Error when fetching venture name.. Check get_name function")
        return "1.xlsx"

    suffix = " - Q" + str(qtr) + venture_info[1]
    prefix = venture_info[0]

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
    print("quarter is " + qtr)
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
    Get venture name from excel sheet using venture name dictionary 
        -Capital Deployment sheet contains venture name in B4
        -AM Fee sheet contains venture name in B6
    Then, rename excel file based on QR venture name

    Currently, this function can handle submissions with multiple ventures within one sheet
    '''
    try:
        wb = xl.load_workbook(filename = file,data_only=True)   
    except:
        print("Probably not an excel file. Ignoring this file...")
        return []

    QR_venture_name = ""
    deployment = False
    am_fee = False
    am_ws = ""

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
                print(venture_name)
                am_fee = True
                am_ws = ws.title
                # print("The sheet with AM Fees is " + am_ws)
            else:
                print("Expected worksheet not found. Check if they submitted with the correct template.")

            qr_name_temp = vars.venture_names.get(venture_name)

            if am_fee:
                if (qr_name_temp != QR_venture_name) & (qr_name_temp is not None):
                    QR_venture_name = qr_name_temp
                    print("QR venture name is " + str(QR_venture_name))
                    fee.recordfee(QR_venture_name,file,am_ws,qtr,year)
                    am_ws = False
                else:
                    print("Venture name not found. Possible that it is empty. Trying next worksheet")

    filename = ""

    if deployment & am_fee:
        filename = " QRI Capital Deployment Forecast and AM Fee.xlsx"
    elif deployment:
        filename = " QRI Capital Deployment Forecast.xlsx"
    elif am_fee:
        filename = " QRI AM Fee.xlsx"
    else:
        print("hmm.. no deployment fee or am fee.. please check the file. ")
        return []

    if QR_venture_name is not None:
        return [QR_venture_name,filename]
    else:
        return []


def exception_ventures():
    print("Venture Exception handling not implemented yet, exiting now...")