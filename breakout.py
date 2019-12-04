# -*- coding: utf-8 -*-
"""
2019
@author: wuffalo

Breaks out groupings from SOS and special lanes
DSLC, Roanoke, RLCA, WWT, IngramMX
"""

import pandas as pd
import xlsxwriter
import os
import glob
from datetime import datetime as dt, timedelta

if os.path.exists("/mnt/c/Users/WMINSKEY/.pen/Breakout_py.xlsx"):
  os.remove("/mnt/c/Users/WMINSKEY/.pen/Breakout_py.xlsx")

test_code = False

def format_sheet():
    worksheet.set_column('A:A',13)
    worksheet.set_column('B:B',45)
    worksheet.set_column('C:C',5)
    worksheet.set_column('D:D',7)
    worksheet.set_column('E:E',22)
    worksheet.set_column('F:F',18)
    worksheet.set_column('G:G',10)
    worksheet.set_column('H:H',4)
    worksheet.set_column('I:I',27)
    worksheet.set_column('J:J',13,format5)
    worksheet.conditional_format('J2:J'+str(len(df.index)+1), {'type': 'duplicate',
                                        'format': format3})
    worksheet.conditional_format('E2:E'+str(len(df.index)+1), {
        'type': 'date',
        'criteria': 'less than',
        'value': (dt.now()-timedelta(1)),
        'format': format1
        })
    worksheet.autofilter('A1:J'+str(len(df.index)+1))

if test_code == True:
    path_to_SOS = "/mnt/c/Users/WMINSKEY/.pen/SOS.csv" # change to latest_file
else:
    list_of_files = glob.glob('/mnt/c/Users/WMINSKEY/Downloads/Shipment Order Summary -*.csv') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    path_to_SOS = latest_file

path_to_excel = "/mnt/c/Users/WMINSKEY/.pen/Breakout_py.xlsx"

show_DSLC = True
show_ROANOKE = True
show_RLCA = True
show_WWT = True
show_IngramMX = True

df = pd.read_csv(path_to_SOS, parse_dates=[11,19], infer_datetime_format=True)

#columns to delete - INITIAL PASS
df = df.drop(columns=['ORDERKEY','SO','SS','STORERKEY','INCOTERMS','ORDERDATE','ACTUALSHIPDATE','DAYSPASTDUE',
                'PASTDUE','ORDERVALUE','TOTALSHIPPED','EXCEP','STOP','PSI_FLAG','UDFNOTES','INTERNATIONALFLAG',
                'BILLING','LOADEDTIME','UDFVALUE1'])

#rename remaining columns
df = df.rename(columns={'EXTERNORDERKEY':'SO-SS','C_COMPANY':'Customer','ADDDATE':'Add Date','STATUSDESCR':'Status',
                        'TOTALORDERED':'QTY','SVCLVL':'Carrier','EXTERNALLOADID':'Load ID','EDITDATE':'Last Edit',
                        'C_STATE':'State','C_COUNTRY':'Country','Textbox6':'TIS'})

writer = pd.ExcelWriter(path_to_excel, engine='xlsxwriter')
workbook = writer.book

# Light red fill with dark red text.
format1 = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})

# Light yellow fill with dark yellow text.
format2 = workbook.add_format({'bg_color':   '#FFEB9C',
                               'font_color': '#9C6500'})

# Green fill with dark green text.
format3 = workbook.add_format({'bg_color':   '#C6EFCE',
                               'font_color': '#006100'})

format5 = workbook.add_format({'num_format': '#'})

#Create DF queries
DSLC = df['TYPEDESCR'] == "DSLC Move"
ROANOKE = df['CUSTID'] == "7128"
RLCA = df['Carrier'] == "RLCA-LTL-4_DAY"
WWT = df['Carrier'] == "TXAP-TL-STD_WWT"
IngramMX = df['Customer'] == "Interamerica Forwarding C/O Ingram Micro Mexi"

#sort table by decreasing importance
df.sort_values(by=['Status','Carrier','Customer','Last Edit','Load ID'], inplace=True)

#drop columns - SECOND PASS
df = df.drop(columns=['TYPEDESCR','CUSTID','PROMISEDATE','Last Edit'])

#Check if dataframes are empty
if DSLC.empty == True:
    show_DSLC = False
if ROANOKE.empty == True:
    show_ROANOKE = False
if RLCA.empty == True:
    show_RLCA = False
if WWT.empty == True:
    show_WWT = False
if IngramMX.empty == True:
    show_IngramMX = False

#Give preview of queries
if test_code == True:
    print("DSLC Orders: \n",df[DSLC].head(2))
    print("Roanoke Orders: \n",df[ROANOKE].head(2))
    print("RLCA Orders: \n",df[RLCA].head(2))
    print("WWT Orders: \n",df[WWT].head(2))
    print("Ingram MX Orders: \n",df[IngramMX].head(2))

#create and format main sheet of all orders
df.to_excel(writer, sheet_name='Main', index=False)
worksheet = writer.sheets['Main']
format_sheet()

#create various sheets if group type is present
if show_DSLC == True:
    df[DSLC].to_excel(writer, sheet_name='DSLC', index=False)
    worksheet = writer.sheets['DSLC']
    format_sheet()
if show_ROANOKE == True:
    df[ROANOKE].to_excel(writer, sheet_name='Roanoke', index=False)
    worksheet = writer.sheets['Roanoke']
    format_sheet()
if show_RLCA == True:
    df[RLCA].to_excel(writer, sheet_name='RLCA', index=False)
    worksheet = writer.sheets['RLCA']
    format_sheet()
if show_WWT == True:
    df[WWT].to_excel(writer, sheet_name='WWT', index=False)
    worksheet = writer.sheets['WWT']
    format_sheet()
if show_IngramMX == True:
    df[IngramMX].to_excel(writer, sheet_name='IngramMX', index=False)
    worksheet = writer.sheets['IngramMX']
    format_sheet()

#color tabs
writer.sheets['Main'].set_tab_color('yellow')
writer.sheets['DSLC'].set_tab_color('green')
writer.sheets['Roanoke'].set_tab_color('orange')
writer.sheets['RLCA'].set_tab_color('red')
writer.sheets['WWT'].set_tab_color('blue')
writer.sheets['IngramMX'].set_tab_color('purple')

writer.save()