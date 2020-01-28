# -*- coding: utf-8 -*-
"""
2019
@author: wuffalo

Breaks out groupings from SOS and special lanes
DSLC, Roanoke, RLCA, WWT, IngramMX, Avt
"""

import pandas as pd
import xlsxwriter
import os
import glob
from datetime import datetime as dt, timedelta

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
    worksheet.autofilter('A1:J'+str(X))

ctime = dt.now()

show_DSLC = True
show_ROANOKE = True
show_RLCA = True
show_WWT = True
show_IngramMX = True
show_Avt = True

output_directory = "/mnt/c/Users/WMINSKEY/.pen/"
output_file_name = "Breakout_py.xlsx"
path_to_output = output_directory+output_file_name

if os.path.exists(path_to_output):
    if os.path.exists(output_directory+'~$'+output_file_name):
        print("File is in use. Close \'"+path_to_output+"\' to try again.")
        raise SystemExit
    else: os.remove(path_to_output)

list_of_files = glob.glob('/mnt/c/Users/WMINSKEY/Downloads/Shipment Order Summary -*.csv') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)
path_to_SOS = latest_file

file_time = os.path.getctime(path_to_SOS)
update_time = dt.fromtimestamp(file_time).strftime('%m/%d/%Y %H:%M')

df = pd.read_csv(path_to_SOS, parse_dates=[11,19], infer_datetime_format=True)

#columns to delete - INITIAL PASS
df = df.drop(columns=['ORDERKEY','SO','SS','STORERKEY','INCOTERMS','ORDERDATE','ACTUALSHIPDATE','DAYSPASTDUE',
                'PASTDUE','ORDERVALUE','TOTALSHIPPED','EXCEP','STOP','PSI_FLAG','UDFNOTES','INTERNATIONALFLAG',
                'BILLING','LOADEDTIME','UDFVALUE1'])

#rename remaining columns
df = df.rename(columns={'EXTERNORDERKEY':'SO-SS','C_COMPANY':'Customer','ADDDATE':'Add Date','STATUSDESCR':'Status',
                        'TOTALORDERED':'QTY','SVCLVL':'Carrier','EXTERNALLOADID':'Load ID','EDITDATE':'Last Edit',
                        'C_STATE':'State','C_COUNTRY':'Country','Textbox6':'TIS'})

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

#Create DF queries
DSLC = df['TYPEDESCR'] == "DSLC Move"
ROANOKE = df['CUSTID'] == "7128"
RLCA = df['Carrier'] == "RLCA-LTL-4_DAY"
WWT = df['Carrier'] == "TXAP-TL-STD_WWT"
IngramMX = df['Customer'] == "Interamerica Forwarding C/O Ingram Micro Mexi"
AVT = df['Carrier'] == "TXAP-TL-STD_MULTISTP"

#find lengths of main dataframe and each query
main_length = len(df.index)
try:
    DSLC_length = df.TYPEDESCR.value_counts()['DSLC Move']
except:
    DSLC_length = 0
try:
    Roanoke_length = df.CUSTID.value_counts()['7128']
except:
    Roanoke_length = 0
try:
    RLCA_length = df.Carrier.value_counts()['RLCA-LTL-4_DAY']
except:
    RLCA_length = 0
try:
    WWT_length = df.Carrier.value_counts()['TXAP-TL-STD_WWT']
except:
    WWT_length = 0
try:
    IngramMX_length = df.Customer.value_counts()['Interamerica Forwarding C/O Ingram Micro Mexi']
except:
    IngramMX_length = 0
try:
    AVT_length = df.Carrier.value_counts()['TXAP-TL-STD_MULTISTP']
except:
    AVT_length = 0

#sort table by decreasing importance
df.sort_values(by=['Status','Carrier','Customer','Last Edit','Load ID'], inplace=True)

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

writer.save()