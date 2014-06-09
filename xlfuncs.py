import xlsxwriter
import pandas as pd
import pdfuncs
import numpy as np
import os
import re

def tab_name(tab):
    return re.sub('[\[\]:*"?/]', '', str(tab))[:30]

def removeNonAscii(s):
    try: return "".join(i for i in s if ord(i)<128)
    except: return s

def cleanDataFrame(df):
    df = df.apply(lambda x: x.fillna(0))
    df = df.apply(lambda x: x.replace([np.inf, -np.inf], 0))
    for col in df:
        if type(df[col].iloc[0]) == str: 
            df[col] = df[col].apply(lambda x: removeNonAscii(x))
    return df

def pivots_to_tabs(df, pivots, filename="output.xlsx", show=True, drop_cols=False):
    # pivots is list columns to group df by 
    df = pdfuncs.cleanDataFrame(df)
    writer = pd.ExcelWriter(filename) 
    if len(pivots)!= 0:
        for p in range(len(pivots)): 
            tbl = df.groupby(pivots[p]).sum()
            tbl = cleanDataFrame(tbl)
            if drop_cols != False: tbl = tbl.drop(drop_cols,1)
            name = tab_name(pivots[p])
            tbl.to_excel(writer,name)
    else:
        df.to_excel(writer,"Data",index=False)    
    writer.save()
    if show == True: os.system("open "+filename)

def dfs_to_tabs(df_list,filename="output.xlsx",show=True):
    # df_list is a list of tuples
    # [(df, tab_name),(df2,tab2_name),...]
    writer = pd.ExcelWriter(filename)
    for i in df_list:
        name = tab_name(i[1])
        tbl = cleanDataFrame(i[0])
        tbl.to_excel(writer,name,index=False)
    writer.save()
    if show == True: os.system("open "+filename)

def custom_tabs(tab_list,filename="output.xslx",finish=True):
    # tab_list is List of Tuples, First Element is a List
    # [([df1, df2], tab_name),([df3], tab_name)] 
    workbook = xlsxwriter.Workbook(filename)
    bold = workbook.add_format({'bold': True})
    for t in range(len(tab_list)): 
        ws = workbook.add_worksheet(tab_name(tab_list[t][1]))
        offset = 0
        tbls = tab_list[t][0]
        for tbl in tbls:
            tbl = cleanDataFrame(tbl)
            headers = tbl.columns
            for h in range(len(headers)): 
                ws.write(offset,h,headers[h],bold)
            for row in range(len(tbl)):
                for col in range(len(tbl.iloc[row])):
                    value = tbl.iloc[row][col]
                    ws.write(offset+row+1,col,value)
            offset = offset + row + 3
    if finish == True: workbook.close();os.system("open "+filename)
    else: return workbook

def create_chart():
    # ...
    return

def format_table():
    # for values for currency, percent, number, etc. 
    # add commas to numbers >= 1M 
    # determine column width appropriately 
    # table headers in bold (extract from custom_tabs func) 
    # align values in cell 
    return






