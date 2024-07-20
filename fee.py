import openpyxl as xl
import pandas as pd
import vars


def add_to_fee_db(df:pd.DataFrame, isqtr:int):
    '''
    append the venture fees to the existing master DataFrame of venture fees
    '''
    if isqtr == 0:
        list = [vars.db,df]
        vars.db = pd.concat(list)
    elif isqtr == -1:
        list = [vars.pq_db,df]
        vars.pq_db = pd.concat(list)
    else:
        list = [vars.nq_db,df]
        vars.nq_db = pd.concat(list)


def recordfee(venture:str, file:str, ws:str, qtr:str, year:str):
    '''
    record the QR venture name and its corresponding AM and Promote Fee on pandas dataframe
    write the dataframe onto Excel sheet, save in Partner Fees folder
    '''
    df = pd.read_excel(file,sheet_name=ws)
    # df = pd.read_excel("C:/Users/harry.ahn/OneDrive - Quadreal Property Group/Portfolio Management/Partner Fees/Capital Deployment Forecast & Actual Fees/2024 Q1/QRI Fee Capital Deployment Forecast Template-04.17.24.xlsx",sheet_name = "Manager Input Fee Payment")
    #this loop is for the corner case where partner adds random rows before the fee rows
    y = 5
    z = 9
    while y < 15:
        try:
            if not pd.isna(df.iloc[y][4]):
                z = y
                break
        except:
            print("No more rows with values on the 5th column")
            break
        y += 1

    df.columns = df.iloc[z]
    # print("The quarters are in row " + str(z+2))

    #this loop is for the corner case where partner uses old template with fee on wrong row
    x = 10
    # rows = [10,11,15,16]
    rows = []
    while len(rows) < 2:
        #start at 10
        try:
            if len(rows) == 0:
                if (df.iloc[x][0] == "EARNED AM FEE FOR THE QUARTER") or \
                    (df.iloc[x][0] == "REALIZED PROMOTE DURING THE QUARTER"):
                    rows.append(x)
            elif not pd.isna(df.iloc[x][0]):
                rows.append(x)
        except:
            print("No more rows with values")
            break
        x += 1
        if x > 20:
            break
    df1 = df.iloc[rows, :5]
    df1.head()
    df1.columns = ['Type', df1.columns[1], df1.columns[2], df1.columns[3]]
    df1.set_index(df1.columns[0])

    df2 = df1.transpose()
    df2.index.name = df1.columns[0] #set index name to 'Type'
    df2.head()
    df2.columns = df2.iloc[0]
    df2 = df2.tail(-1)

    # qtr = "1"
    # year = "2024"
    str_match = "Q" + qtr + " " + year
    prev_qtr_match = getprevqtr(qtr, year)
    next_qtr_match = getnextqtr(qtr, year)
    # prev_qtr_match = "Q1 2023"

    try:
        df_curr = df2[df2.index.str.match(str_match)]
        df_curr.head()
        # venture = "Charter Hall"
        df_curr = df_curr.rename(index={str_match : venture})
        df_curr_f = fillemptyfee(df_curr)
        df_curr_f.head()
        df_curr_f_abs = df_curr_f.abs()
        print(df_curr_f_abs.iloc[0][0])
        add_to_fee_db(df_curr_f_abs,0)
    except:
        print("Current quarter data not found. Skipping fee collection...")
    
    try:
        df_prev = df2[df2.index.str.match(prev_qtr_match)]
        df_prev.head()
        df_prev = df_prev.rename(index={prev_qtr_match : venture})
        df_prev_f = fillemptyfee(df_prev)
        df_prev_f_abs = df_prev_f.abs()
        add_to_fee_db(df_prev_f_abs,-1)
    except:
        print("Previous quarter data not found.")

    try:
        df_next = df2[df2.index.str.match(next_qtr_match)]
        df_next.head()
        df_next = df_next.rename(index={next_qtr_match : venture})
        df_next_f = fillemptyfee(df_next)
        df_next_f_abs = df_next_f.abs()
        add_to_fee_db(df_next_f_abs,1)
    except:
        print("Next quarter data not found.")

    print(str(venture) + "done")

    
def fillemptyfee(df:pd.DataFrame)->pd.DataFrame:
    '''
    This loop is for the corner case where partner deleted unused fee rows (ex. accrued promote)
    Fill those rows with value 0, 
    '''
    columns = ['EARNED AM','REALIZED PROMOTE']
    if len(df.columns) < len(columns):
        n = len(columns) - len(df.columns)
        while n < len(columns):
            df[str(n)] = 0
            n += 1
    df.columns = columns
    return df


def getprevqtr(qtr: str, year: str)->str:
    prev_qtr = 0
    prev_qtr_year = 0

    if (int(qtr) == 1):
        prev_qtr = "4"
        prev_qtr_year = str(int(year) - 1)
    else:
        prev_qtr = str(int(qtr) - 1)
        prev_qtr_year = year

    output = "Q" + prev_qtr + " " + prev_qtr_year
    # print("Previous quarter is " + output)
    return output


def getnextqtr(qtr: str,year: str)->str:
    next_qtr = 0
    next_qtr_year = 0

    if (int(qtr) == 4):
        next_qtr = "1"
        next_qtr_year = str(int(year)+1)
    else:
        next_qtr = str(int(qtr) + 1)
        next_qtr_year = year

    output = "Q" + next_qtr + " " + next_qtr_year
    # print("Next quarter is " + output)
    return output


def export_fee_db():
    '''
    export master DataFrame
    '''
    # print(vars.db)
    final_excel = vars.excel
    with pd.ExcelWriter(
        final_excel,
        engine='openpyxl',
        mode='a',
        if_sheet_exists='replace'
    ) as writer:
        vars.db.to_excel(writer, sheet_name="Current Qtr")
        vars.pq_db.to_excel(writer,sheet_name="Prev Qtr")
        vars.nq_db.to_excel(writer,sheet_name="Next Qtr")