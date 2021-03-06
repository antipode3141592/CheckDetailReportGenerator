#   XLSX Report Generator
#   Copyright 2018 by Sean Vo Kirkpatrick using GNU GPL v3
#   skirkpatrick@racc.org or sean@studioantipode.com or seanvokirkpatrick@gmail.com
#   
#   Create a .xlsx report for each fund
#
#   Tested using    - Anaconda 5.0.0
#                   - pandas 0.22.0
#                   - XlsxWriter 1.0.5
#                   - pyodbc 4.0.23
#   IDE: Visual Studio 2017 Community Edition

# License Info:
#   This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.

import datetime as dt
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_range
import win32com.client as win32
import pyodbc 
import os
import numpy as np

# define those functions!
def writerow(ws,r,i,row,width):
    addresslines = "%s, %s, %s %s" %(str(row['Fund_Address']), str(row['Fund_City']), str(row['Fund_State']), str(row['Fund_Zip'],))
    #print(addresslines)
    if len(str(row['Fund'])) > width[0]:
        width[0] = len(str(row['Fund']))
    if len(str(addresslines)) > width[1]:
        width[1] = len(str(addresslines))
    if len(str(row['Appeal'])) > width[2]:
        width[2] = len(str(row['Appeal']))
    if len(str(row['GiftDate'])) > width[3]:
        width[3] = len(str(row['GiftDate']))
    if len(str(row['Name'])) > width[4]:
        width[4] = len(str(row['Name']))
    if len(str(row['Reference'])) > width[5]:
        width[5] = len(str(row['Reference']))
    if len(str(row['SplitAmount'])) > width[6]:
        width[6] = len(str(row['SplitAmount']))
    ws.write(r,0,row['Fund'])
    ws.write(r,1,addresslines)
    ws.write(r,2,row['Appeal'])
    ws.write(r,3,row['GiftDate'], fmt_date)
    ws.write(r,4,row['Name'])
    ws.write(r,5,row['Reference'])
    ws.write(r,6,row['SplitAmount'], fmt_money)
    r+=1
    return r, width

def writerow2(ws,r,i,row,width):
    #print(row)
    if len(str(row['GiftDate'])) > width[0]:
        width[0] = len(str(row['GiftDate']))
    if len(str(row['Name'])) > width[1]:
        width[1] = len(str(row['Name']))
    if len(str(row['Reference'])) > width[2]:
        width[2] = len(str(row['Reference']))
    if len(str(row['Reference'])) > width[3]:
        width[3] = len(str(row['GiftID']))
    if len(str(row['SplitAmount'])) > width[4]:
        width[4] = len(str(row['SplitAmount']))
    ws.write(r,0,row['GiftDate'], fmt_date)
    ws.write(r,1,row['Name'])
    ws.write(r,2,row['Reference'])
    ws.write(r,3,row['GiftID'])
    ws.write(r,4,row['SplitAmount'], fmt_money)
    r+=1
    return r, width

def writetotal(ws,fmt,r,row_1st,row_last):
    ws.write(r,0,"Total",fmt)
    ws.write(r,1,"", fmt)
    ws.write(r,2,"", fmt)
    ws.write(r,3,"", fmt)
    formula = "=SUBTOTAL(109,{:s})".format(xl_range(row_1st,4,row_last,4))
    ws.write_formula(xl_rowcol_to_cell(r,4), formula, fmt)
    return

def writeevalrow(ws,r,i,row,width):
    if len(str(i[1])) > width[0]:
        width[0] = len(str(i[1]))
    ws.write(r,0,i[1])
    ws.write(r,1,row['SplitAmount'],fmt_money)
    if (row['SplitAmount'] >= 20):
        ws.write(r,2,"Y")
    else:
        ws.write(r,2,"N")
    r+=1
    return r, width

def writemergerow(ws,r,k,df,width,total):
    #columns: ['FUND_ID', 'ORG_NAME', 'Category', 'ADDRESS_BLOCK', 'CITY', 'STATE', 'POST_CODE']
    address = df.loc[k[0]][2]
    city = df.loc[k[0]][3]
    state = df.loc[k[0]][4]
    zipcode = df.loc[k[0]][5]
    citystatezip = "%s, %s %s" %(str(city), str(state), str(zipcode))
    if len(str(k[1])) > width[0]:
        width[0] = len(str(k[1]))
    if len(str(address)) > width[2]:
        width[2] = len(str(address))
    if len(str(citystatezip)) > width[3]:
        width[3] = len(str(citystatezip))
    ws.write(r,0,k[1])
    ws.write(r,1,total,fmt_money)
    ws.write(r,2,address)
    ws.write(r,3,citystatezip)
    r+=1
    return r, width

def sterilizestring (s):
    for char in "?.!/;:":
        s = s.replace(char,'_');
    return s

print("Processing...")
cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};" #requires explicitily stating the sql driver
                      "Server=overlook;"
                      "Database=re_racc;"
                      "Trusted_Connection=yes;")    #use windows integrated security
cursor = cnxn.cursor()

#-----------------------------------------------------------------
#inputs
startdate = '2019-12-01'    #YYYY-DD-MM format
enddate = '2019-12-31'
#-----------------------------------------------------------------


sqlcommand = 'exec sp_GiftReconwithAddress ''?'', ''?'''
sqlcommand2 = 'exec sp_OrgswithAddresses'
sqlparams = (startdate,enddate)

cursor.execute(sqlcommand,sqlparams)
columns = [column[0] for column in cursor.description]
#['Campaign', 'Fund', 'Fund ID', 'Fund Address', 'Fund City', 'Fund State', 'Fund Zip', 'Appeal', 'GiftID', 'RECORDS_ID', 'Type', 'GiftType', 'TotalAmount', 'SplitAmount', 'AnonRecord', 'AnonGift', 'CONSTITUENT_ID', 'FIRST_NAME', 'KEY_NAME', 'Ref', 'Name', 'Reference', 'POST_DATE', 'CHECK_DATE', 'CHECK_NUMBER', 'GiftDate', 'BATCH_NUMBER', 'FundCategory', 'MGConstituentID']
print(columns)
data = []   #grab results, put into a list, put list into numpy array, and then put numpy array into pandas dataframe
for row in cursor:
    data.append(tuple(row))
df_report = pd.DataFrame.from_records(np.array(data),columns=columns)

cursor.execute(sqlcommand2)
columns = [column[0] for column in cursor.description]
data = []
for row in cursor:
    data.append(tuple(row))
df_addresses = pd.DataFrame.from_records(np.array(data),columns=columns)

filepath = "C:\\Users\\skirkpatrick\\Coding\\Python\\"
outputpath = "C:\\Users\\skirkpatrick\\Coding\\Python\\Outgoing\\"
fl = filepath + "Check Request - Quarter end 12-31-2019.xlsx"

#create primary DG payment report (single file)
with xlsxwriter.Workbook(fl, {'nan_inf_to_errors': True}) as wb:
    #define formats
    fmt_money = wb.add_format({'num_format': '$#,##0.00'})
    fmt_date = wb.add_format({'num_format': 'mm/dd/yyyy'})
    fmt_dataheader = wb.add_format({'bold': True, 'bg_color': '#000000', 'font_color': '#FFFFFF' })
    fmt_total = wb.add_format({'bg_color': '#434343', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1, 'font_color': '#FFFFFF'})   #darkest grey, white bold text

    #-------------------------------------------------------------------
    #Hold Evaluation sheet
    ws = wb.add_worksheet("Hold Eval")
    ws.set_landscape()
    ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
    startingdatarow = 1     #indicates which row to start writing data to
    ws.repeat_rows(startingdatarow-1) #repeats header row on each page for printing (r-1 because it uses excel row numbers, not 0-index rows)
    #write header
    ws.write(startingdatarow-1,0,'Payable To',fmt_dataheader)
    ws.write(startingdatarow-1,1,'Total',fmt_dataheader)
    ws.write(startingdatarow-1,2,'Pay Out?',fmt_dataheader)
    r = startingdatarow
    widths = [10,10,10]
    sumsoffunds = df_report.groupby(['FUND_ID','Fund']).agg({'SplitAmount':'sum'})
    for k,row in sumsoffunds.iterrows():
        r,widths = writeevalrow(ws,r,k,row,widths)
    ws.set_column(0,0,widths[0])
    ws.set_column(1,1,widths[1])
    ws.set_column(2,2,widths[2])
    ws.add_table(0,0,r-1,2,{'columns':[{'header':'Payable To'},
                                       {'header':'Total'},
                                       {'header':'Pay Out?'}]})

    #-------------------------------------------------------------------
    #Check Request Worksheet and Hold Worksheet
    # currently, Check Request sheet must have one last manual step, selecting the whole table (data + headers) and then Data->Subtotal sum by Fund Split Amount
    # 
    header1 = "&LAIF Designated Gift Check Request" + "&CIncludes gifts with dates between " + startdate + " and " + enddate
    footer1 = "&LCoding: 01-5210-Other-280-0-0-0" + "&RApproved as per AIF Designated Gifts for November thru December 2019"
    ws = wb.add_worksheet("Check Request");
    ws.set_header(header1)
    ws.set_footer(footer1)
    ws.hide_gridlines(0)
    ws.set_landscape()
    ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
    
    startingdatarow = 1     #indicates which row to start writing data to
    ws.repeat_rows(startingdatarow-1) #repeats header row on each page for printing (r-1 because it uses excel row numbers, not 0-index rows)
    #write header
    ws.write(startingdatarow-1,0,'Fund',fmt_dataheader)
    ws.write(startingdatarow-1,1,'Fund Address',fmt_dataheader)
    ws.write(startingdatarow-1,2,'Appeal ID',fmt_dataheader)
    ws.write(startingdatarow-1,3,'Gift Date',fmt_dataheader)
    ws.write(startingdatarow-1,4,'Name',fmt_dataheader)
    ws.write(startingdatarow-1,5,'Reference',fmt_dataheader)
    ws.write(startingdatarow-1,6,'Fund Split Amount',fmt_dataheader)

    ws_hold = wb.add_worksheet("Holds")
    ws_hold.write(startingdatarow-1,0,'Fund',fmt_dataheader)
    ws_hold.write(startingdatarow-1,1,'Fund Address',fmt_dataheader)
    ws_hold.write(startingdatarow-1,2,'Appeal ID',fmt_dataheader)
    ws_hold.write(startingdatarow-1,3,'Gift Date',fmt_dataheader)
    ws_hold.write(startingdatarow-1,4,'Name',fmt_dataheader)
    ws_hold.write(startingdatarow-1,5,'Reference',fmt_dataheader)
    ws_hold.write(startingdatarow-1,6,'Fund Split Amount',fmt_dataheader)
    
    r = startingdatarow
    r_hold = startingdatarow
    widths = [10,10,10,10,10,10,10]
    widths_hold = [10,10,10,10,10,10,10]
    for k, row in df_report.iterrows():
        #if sumsoffunds['SplitAmount'][row['FUND_ID']].values[0] >= 20:
        r, widths = writerow(ws,r,k,row,widths)
        #else:
        #    r_hold, widths = writerow(ws_hold,r_hold,k,row,widths)
    ws.set_column(0,0,widths[0])
    ws.set_column(1,1,widths[1])
    ws.set_column(2,2,widths[2])
    ws.set_column(3,3,widths[3])
    ws.set_column(4,4,widths[4])
    ws.set_column(5,5,widths[5])
    ws.set_column(6,6,widths[6])
    ws_hold.set_column(0,0,widths[0])
    ws_hold.set_column(1,1,widths[1])
    ws_hold.set_column(2,2,widths[2])
    ws_hold.set_column(3,3,widths[3])
    ws_hold.set_column(4,4,widths[4])
    ws_hold.set_column(5,5,widths[5])
    ws_hold.set_column(6,6,widths[6])


    #-------------------------------------------------------------------
    #Mail Merge
    # currently, Check Request sheet must have one last manual step, selecting the whole table (data + headers) and then Data->Subtotal sum by Fund Split Amount
    # 
    header1 = "&LAIF Designated Gift Check Request" + "&CIncludes gifts with dates between " + startdate + " and " + enddate
    footer1 = "&LCoding: 01-5210-Other-280-0-0-0" + "&RApproved as per AIF Designated Gifts for November thru December 2019"
    ws = wb.add_worksheet("Mail Merge");
    ws.set_header(header1)
    ws.set_footer(footer1)
    ws.hide_gridlines(0)
    ws.set_landscape()
    ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
    
    startingdatarow = 1     #indicates which row to start writing data to
    ws.repeat_rows(startingdatarow-1) #repeats header row on each page for printing (r-1 because it uses excel row numbers, not 0-index rows)
    #write header
    ws.write(startingdatarow-1,0,'Payable_to',fmt_dataheader)
    ws.write(startingdatarow-1,1,'Cash',fmt_dataheader)
    ws.write(startingdatarow-1,2,'Fund_Address',fmt_dataheader)
    ws.write(startingdatarow-1,3,'Fund_City_State_ZIP',fmt_dataheader)  
    r = startingdatarow
    widths = [10,10,10,10]
    #columns:  ['FUND_ID', 'ORG_NAME', 'Category', 'ADDRESS_BLOCK', 'CITY', 'STATE', 'POST_CODE']
    _df2 = df_addresses.set_index('FUND_ID')
    for k, row in sumsoffunds.iterrows():
        if ((k[0] in _df2.index) & (row['SplitAmount'] >= 20)):
            r, widths = writemergerow(ws,r,k,_df2,widths,row['SplitAmount'])
    ws.set_column(0,0,widths[0])
    ws.set_column(1,1,widths[1])
    ws.set_column(2,2,widths[2])
    ws.set_column(3,3,widths[3])
    ws_hold.set_column(0,0,widths[0])
    ws_hold.set_column(1,1,widths[1])
    ws_hold.set_column(2,2,widths[2])
    ws_hold.set_column(3,3,widths[3])

#create individual reports for each fund_id, with title of format {Fund_ID}.xlxs
df3 = df_report.groupby(['FUND_ID','Fund'])
for group, data in df3:
    fl2 = outputpath + sterilizestring(group[0]) + ".xlsx"
    print("Writing to file %s..." %(fl2))
    #print(data)
    with xlsxwriter.Workbook(fl2, {'nan_inf_to_errors': True}) as wb:
        #individual fund report
        fmt_money = wb.add_format({'num_format': '$#,##0.00'})  #add formats
        fmt_date = wb.add_format({'num_format': 'mm/dd/yyyy'})
        fmt_dataheader = wb.add_format({'bold': True, 'bg_color': '#000000', 'font_color': '#FFFFFF' })
        fmt_total = wb.add_format({'bg_color': '#434343', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1, 'font_color': '#FFFFFF'})   #darkest grey, white bold text

        header2 = "&LDesignated Gift Payment Details" + "&RIncludes gifts received between " + startdate + " and " + enddate
        footer2 = "&LDesignated Gift Payments to %s" %(group[1]) + "&RGiven via Arts Impact Fund at RACC"
        ws = wb.add_worksheet("Gift Payments");
        ws.set_header(header2)
        ws.set_footer(footer2)
        ws.hide_gridlines(0)
        ws.set_landscape()
        ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
    
        startingdatarow = 1     #indicates which row to start writing data to
        ws.repeat_rows(startingdatarow-1) #repeats header row on each page for printing (r-1 because it uses excel row numbers, not 0-index rows)
        #write header
        ws.write(startingdatarow-1,0,'Gift Date',fmt_dataheader)
        ws.write(startingdatarow-1,1,'Donor Name',fmt_dataheader)
        ws.write(startingdatarow-1,2,'Reference',fmt_dataheader)
        ws.write(startingdatarow-1,3,'Gift ID',fmt_dataheader)
        ws.write(startingdatarow-1,4,'Amount',fmt_dataheader)
        r = startingdatarow
        r_hold = startingdatarow
        widths = [10,10,10,10,10]
        widths_hold = [10,10,10,10,10]
        #print(data)
        for k, row in data.iterrows():
            r, widths = writerow2(ws,r,group,row,widths)
        writetotal(ws, fmt_total, r, startingdatarow,r-1)
        ws.set_column(0,0,widths[0])
        ws.set_column(1,1,widths[1])
        ws.set_column(2,2,widths[2])
        ws.set_column(3,3,widths[3])
        wb.close()
print("Done!")