# -*- coding: utf-8 -*-
"""
2019
@author: wuffalo

Breaks out groupings from SOS and special lanes
DSLC, Roanoke, RLCA, WWT, IngramMX, Avt
"""

import pandas as pd
#import xlsxwriter # included in pandas
import os
import glob
from datetime import datetime as dt, timedelta
#import pandas.io.formats.excel

def format_sheet(X):
    X = X+1
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

ctime = dt.now()

# initializing block showing which filters are actively in use for corresponding output sheets
show_DSLC = True
show_ROANOKE = True
show_RLCA = True
show_WWT = True
show_IngramMX = True
show_Avt = True
show_Rockwell = True

output_directory = "/mnt/c/Users/WMINSKEY/.pen/"
output_file_name = "Breakout_py.xlsx"
path_to_output = output_directory+output_file_name

# finds previous output file and quits program while notifying if file is already open. Otherwise removes old output file.
if os.path.exists(path_to_output):
    if os.path.exists(output_directory+'~$'+output_file_name):
        print("File is in use. Close \'"+path_to_output+"\' to try again.")
        raise SystemExit
    else: os.remove(path_to_output)

list_of_files = glob.glob('/mnt/c/Users/WMINSKEY/Downloads/Shipment Order Summary -*.csv') # * means all if need specific format then *.csv
path_to_SOS = latest_file = max(list_of_files, key=os.path.getctime)

file_time = os.path.getctime(path_to_SOS)
update_time = dt.fromtimestamp(file_time).strftime('%m/%d/%Y %H:%M')

df = pd.read_csv(path_to_SOS, parse_dates=[11,19], infer_datetime_format=True)

#columns to delete - INITIAL PASS
df = df.drop(columns=['ORDERKEY','SO','SS','STORERKEY','INCOTERMS','ORDERDATE','ACTUALSHIPDATE','DAYSPASTDUE',
                'PASTDUE','ORDERVALUE','TOTALSHIPPED','EXCEP','STOP','PSI_FLAG','SUSR5','INTERNATIONALFLAG',
                'LOADEDTIME','UDFVALUE1','ROUTE'])

#rename remaining columns
df = df.rename(columns={'EXTERNORDERKEY':'SO-SS','C_COMPANY':'Customer','ADDDATE':'Add Date','STATUSDESCR':'Status',
                        'TOTALORDERED':'QTY','SVCLVL':'Carrier','EXTERNALLOADID':'Load ID','EDITDATE':'Last Edit',
                        'C_STATE':'State','C_COUNTRY':'Country','Textbox6':'TIS','BILLING':'Route'})

#remove commas from number columns, allows for reading as number then formatting on output
# df['QTY'] = df['QTY'].str.replace(',', '')

#create xlsxwriter object
writer = pd.ExcelWriter(path_to_output, engine='xlsxwriter', options={'strings_to_numbers': True})
workbook = writer.book

# Light red fill with dark red text.
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
format7 = workbook.add_format({'align': 'left'})
#format7 = workbook.add_format()
#format7 = format7.set_align('left')

#Create DF queries, these are boolean masks
DSLC = df['TYPEDESCR'] == "DSLC Move"
ROANOKE = df['CUSTID'] == "7128"
RLCA = df['Carrier'] == "RLCA-LTL-4_DAY"
WWT = df['Carrier'] == "TXAP-TL-STD_WWT"
IngramMX = df['Customer'] == "Interamerica Forwarding C/O Ingram Micro Mexi"
AVT = df['CUSTID'] == "401778414"
ROCK = (df['CUSTID'] == '68275') & (df['State'] == 'IN')

#find lengths of main dataframe and each query, null causes default 0 assignment
main_length = len(df.index)
try:
    DSLC_length = sum(DSLC)
except:
    DSLC_length = 0
try:
    Roanoke_length = sum(ROANOKE)
except:
    Roanoke_length = 0
try:
    RLCA_length = sum(RLCA)
except:
    RLCA_length = 0
try:
    WWT_length = sum(WWT)
except:
    WWT_length = 0
try:
    IngramMX_length = sum(IngramMX)
except:
    IngramMX_length = 0
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

#create pivot table queries
gen_table = pd.pivot_table(df, index=['Carrier','Status'], values=['QTY','Last Edit'], aggfunc={'QTY':'sum'},margins=False)
gen_summary = pd.pivot_table(df, index=['Status'], values=['SO-SS','QTY'], aggfunc={'SO-SS':len,'QTY':'sum'}, margins=False)

#drop columns - SECOND PASS after calculations are performed
df = df.drop(columns=['TYPEDESCR','CUSTID','PROMISEDATE','Last Edit'])

#Check if dataframes are empty
if DSLC_length == 0:
    show_DSLC = False
if Roanoke_length == 0:
    show_ROANOKE = False
if RLCA_length == 0:
    show_RLCA = False
if WWT_length == 0:
    show_WWT = False
if IngramMX_length == 0:
    show_IngramMX = False
if AVT_length == 0:
    show_Avt = False
if ROCK_length == 0:
    show_Rockwell = False

#create and format main sheet of all orders
df.to_excel(writer, sheet_name='Main', index=False)
worksheet = writer.sheets['Main']
format_sheet(main_length)
writer.sheets['Main'].set_tab_color('yellow')
worksheet.write('M1',"Last Update at: "+str(update_time))

#create various sheets if group type is present
if show_DSLC == True:
    df.loc[DSLC].to_excel(writer, sheet_name='DSLC', index=False)
    worksheet = writer.sheets['DSLC']
    format_sheet(DSLC_length)
    writer.sheets['DSLC'].set_tab_color('green')
if show_ROANOKE == True:
    df.loc[ROANOKE].to_excel(writer, sheet_name='Roanoke', index=False)
    worksheet = writer.sheets['Roanoke']
    format_sheet(Roanoke_length)
    writer.sheets['Roanoke'].set_tab_color('orange')
if show_RLCA == True:
    df.loc[RLCA].to_excel(writer, sheet_name='RLCA', index=False)
    worksheet = writer.sheets['RLCA']
    format_sheet(RLCA_length)
    writer.sheets['RLCA'].set_tab_color('red')
if show_WWT == True:
    df.loc[WWT].to_excel(writer, sheet_name='WWT', index=False)
    worksheet = writer.sheets['WWT']
    format_sheet(WWT_length)
    writer.sheets['WWT'].set_tab_color('blue')
if show_IngramMX == True:
    df.loc[IngramMX].to_excel(writer, sheet_name='IngramMX', index=False)
    worksheet = writer.sheets['IngramMX']
    format_sheet(IngramMX_length)
    writer.sheets['IngramMX'].set_tab_color('purple')
if show_Avt == True:
    df.loc[AVT].to_excel(writer, sheet_name='Avt', index=False)
    worksheet = writer.sheets['Avt']
    format_sheet(AVT_length)
    writer.sheets['Avt'].set_tab_color('#33CCCC')
if show_Rockwell == True:
    df.loc[ROCK].to_excel(writer, sheet_name='Rockwell', index=False)
    worksheet = writer.sheets['Rockwell']
    format_sheet(ROCK_length)
    writer.sheets['Rockwell'].set_tab_color('purple')

to_allocate = gen_table.query('Status == ["Allocated","Created Externally"]')

try: # add try/exception block because no orders in Allocated or Created External causes crash
    to_allocate.to_excel(writer, sheet_name='To Start')
    worksheet = writer.sheets['To Start']
    worksheet.set_column('A:A',30,format7)
    worksheet.set_column('B:B',20)
    worksheet.set_column('C:C',6,format5)
except:
    print("No orders in Allocated or Created External. 'To Start' sheet not populated.")

review = gen_summary.query('Status == ["Allocated","Created Externally","In Picking","QA Complete","Pack Ready","Released","Loaded"]')

review.to_excel(writer, sheet_name='Summary')
worksheet = writer.sheets['Summary']
worksheet.set_column('A:A',20,format7)
worksheet.set_column('B:B',10)
worksheet.set_column('C:C',10)

writer.save()