# -*- coding: utf-8 -*-
"""
2021
@author: wuffalo

Breaks out groupings from SOS and special lanes
DSLC, Roanoke, RLCA, WWT, IngramMX, Avt
"""

import pandas as pd
import os
import glob
from datetime import datetime as dt, timedelta

def format_sheet(X,workbook,worksheet):
    format1 = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})
    # orange fill with dark orange text.
    format2 = workbook.add_format({'bg_color':   '#ffcc99',
                                'font_color': '#804000'})
    # yellow fill with dark yellow text.
    format3 = workbook.add_format({'bg_color':    '#ffeb99',
                                    'font_color':   '#806600'})
    # Green fill with dark green text.
    format4 = workbook.add_format({'bg_color':   '#C6EFCE',
                                'font_color': '#006100'})
    format5 = workbook.add_format({'num_format': '#'})
    format6 = workbook.add_format({'num_format': '#,##0'})

    X = X+1
    ctime = dt.now()

    worksheet.set_column('A:A',13)
    worksheet.set_column('B:B',45)
    worksheet.set_column('C:C',7)
    worksheet.set_column('D:D',9)
    worksheet.set_column('E:E',19)
    worksheet.set_column('F:F',18)
    worksheet.set_column('G:G',10)
    worksheet.set_column('H:H',7,format6)
    worksheet.set_column('I:I',29)
    worksheet.set_column('J:J',13,format5)
    worksheet.conditional_format('J2:J'+str(X), {'type': 'duplicate',
                                        'format': format4})
    worksheet.conditional_format('E2:E'+str(X), {
        'type': 'date',
        'criteria': 'less than or equal to',
        'value': (ctime-timedelta(1)),
        'format': format1
        })
    worksheet.conditional_format('E2:E'+str(X), {
        'type': 'date',
        'criteria': 'between',
        'minimum': ctime-timedelta(11/12),
        'maximum': ctime-timedelta(1),
        'format': format2
        })
    worksheet.conditional_format('E2:E'+str(X), {
        'type': 'date',
        'criteria': 'between',
        'minimum': ctime-timedelta(4/5),
        'maximum': ctime-timedelta(11/12),
        'format': format3
        })
    worksheet.autofilter('A1:K'+str(X))

def create_sheet(path,DFsearch,tabname,length,tabcolor):
    writer = pd.ExcelWriter(path, engine='xlsxwriter', options={'strings_to_numbers': True})
    workbook = writer.book
    df.loc[DFsearch].to_excel(writer, sheet_name=tabname, index=False)
    worksheet = writer.sheets[tabname]
    format_sheet(length,workbook,worksheet)
    writer.sheets[tabname].set_tab_color(tabcolor)
    worksheet.write('M1',"Last Update at: "+str(update_time))
    writer.save()

def find_readable_update_time(x): #x is full path to file you want a formatted time value of
    file_time = os.path.getctime(x)
    return dt.fromtimestamp(file_time).strftime('%m/%d/%Y %H:%M')

# initializing block showing which filters are actively in use for corresponding output sheets
show_DSLC = True
show_ROANOKE = True
show_WWT = True
show_Avt = True
show_Rockwell = True

output_directory = "/mnt/c/Users/WMINSKEY/.pen/Breakouts/"
output_file_name = "Breakout_py.xlsx"
path_to_output = output_directory+output_file_name
DSLC_file_name = "DSLC.xlsx"
path_to_DSLC = output_directory+DSLC_file_name
ROAN_file_name = "Roanoke.xlsx"
path_to_ROAN = output_directory+ROAN_file_name
WWT_file_name = "WWT.xlsx"
path_to_WWT = output_directory+WWT_file_name
AVT_file_name = "Avt.xlsx"
path_to_Avt = output_directory+AVT_file_name
Rock_file_name = "Rockwell.xlsx"
path_to_Rock = output_directory+Rock_file_name

list_of_files = glob.glob('/mnt/c/Users/WMINSKEY/Downloads/Shipment Order Summary -*.csv') # * means all if need specific format then *.csv
path_to_SOS = latest_file = max(list_of_files, key=os.path.getctime)
update_time = find_readable_update_time(path_to_SOS)

df = pd.read_csv(path_to_SOS, parse_dates=[11,19], infer_datetime_format=True)

#columns to delete - INITIAL PASS
df = df.drop(columns=['ORDERKEY','SO','SS','STORERKEY','INCOTERMS','ORDERDATE','ACTUALSHIPDATE','DAYSPASTDUE',
                'PASTDUE','ORDERVALUE','TOTALSHIPPED','EXCEP','STOP','PSI_FLAG','SUSR5','INTERNATIONALFLAG',
                'LOADEDTIME','UDFVALUE1','ROUTE'])

#rename remaining columns
df = df.rename(columns={'EXTERNORDERKEY':'SO-SS','C_COMPANY':'Customer','ADDDATE':'Add Date','STATUSDESCR':'Status',
                        'TOTALORDERED':'QTY','SVCLVL':'Carrier','EXTERNALLOADID':'Load ID','EDITDATE':'Last Edit',
                        'C_STATE':'State','C_COUNTRY':'Country','Textbox6':'TIS','BILLING':'Route'})

#Create DF queries, these are boolean masks
DSLC = df['TYPEDESCR'] == "DSLC Move"
ROANOKE = df['CUSTID'] == "7128"
WWT = df['Carrier'] == "TXAP-TL-STD_WWT"
AVT = df['CUSTID'] == "401778414"
ROCK = (df['CUSTID'] == '68275') & (df['State'] == 'IN')

#find lengths of main dataframe and each query, null causes default 0 assignment
try:
    DSLC_length = sum(DSLC)
except:
    DSLC_length = 0
try:
    Roanoke_length = sum(ROANOKE)
except:
    Roanoke_length = 0
try:
    WWT_length = sum(WWT)
except:
    WWT_length = 0
try:
    AVT_length = sum(AVT)
except:
    AVT_length = 0
try:
    ROCK_length = sum(ROCK)
except:
    ROCK_length = 0

#sort table by decreasing importance
df.sort_values(by=['Status','Carrier','Customer','Last Edit','Load ID'], inplace=True)

#drop columns - SECOND PASS after calculations are performed
df = df.drop(columns=['TYPEDESCR','CUSTID','PROMISEDATE','Last Edit'])

#Check if dataframes are empty
if DSLC_length == 0:
    show_DSLC = False
if Roanoke_length == 0:
    show_ROANOKE = False
if WWT_length == 0:
    show_WWT = False
if AVT_length == 0:
    show_Avt = False
if ROCK_length == 0:
    show_Rockwell = False

if show_DSLC == True:
    try:
        create_sheet(path_to_DSLC,DSLC,'DSLC',DSLC_length,'green')
    except:
        print("The DSLC file is open. Cannot update.")
if show_ROANOKE == True:
    try:
        create_sheet(path_to_ROAN,ROANOKE,'Roanoke',Roanoke_length,'orange')
    except:
        print("The Roanoke file is open. Cannot update.")
if show_WWT == True:
    try:
        create_sheet(path_to_WWT,WWT,'WWT',WWT_length,'blue')
    except:
        print("The WWT file is open. Cannot update.")
if show_Avt == True:
    try:
        create_sheet(path_to_Avt,AVT,'Avt',AVT_length,'#33CCCC')
    except:
        print("The Avt file is open. Cannot update.")
if show_Rockwell == True:
    try:
        create_sheet(path_to_Rock,ROCK,'Rockwell',ROCK_length,'purple')
    except:
        print("The Rockwell file is open. Cannot update.")