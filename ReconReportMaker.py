#   XLSX Report Generator
#   Copyright 2018 by Sean Vo Kirkpatrick using GNU GPL v3
#   skirkpatrick@racc.org or sean@studioantipode.com or seanvokirkpatrick@gmail.com
#   
#   Creates a formatted Excel .xlsx summary report with subtotals from each batch in an input file
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
def writeYTDsummary_appeal(ws,r,i,row,width):
    if len(str(i)) > width[0]:
        width[0] = len(str(i))
    if len(str(row.values[0])) > width[1]:
        width[1] = len(str(row.values[0]))
    ws.write(r,0,i)
    ws.write(r,1,row.values[0],fmt_money)
    r+=1
    return r, width

def writeYTDsummary_appealfund(ws,r,i,row,width):
    if len(str(i[0])) > width[0]:
        width[0] = len(str(i[0]))
    if len(str(i[1])) > width[1]:
        width[1] = len(str(i[1]))
    if len(str(row.values[0])) > width[2]:
        width[2] = len(str(row.values[0]))
    ws.write(r,0,i[0])
    ws.write(r,1,i[1])
    ws.write(r,2,row.values[0],fmt_money)
    r+=1
    return r, width

def sterilizestring (s):
    for char in "?.!/;:":
        s = s.replace(char,'_');
    return s
#-----------------------------------------------------------------
print("Processing...")
cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};" #requires explicitily stating the sql driver
                      "Server=overlook;"
                      "Database=re_racc;"
                      "Trusted_Connection=yes;")    #use windows integrated security
cursor = cnxn.cursor()
startdate = '2017-07-01'
enddate = '2018-06-30'
sqlcommand = 'exec sp_giftreconreport ''?'', ''?'''
sqlparams = (startdate,enddate)
cursor.execute(sqlcommand,sqlparams)
columns = [column[0] for column in cursor.description]
#['Campaign', 'Fund', 'Appeal', 'GiftID', 'RECORDS_ID', 'Type', 'GiftType', 'TotalAmount', 'SplitAmount', 'AnonRecord', 'AnonGift', 'CONSTITUENT_ID', 'FIRST_NAME', 'KEY_NAME', 'Ref', 'Name', 'Reference', 'POST_DATE', 'CHECK_DATE', 'CHECK_NUMBER', 'GiftDate', 'BATCH_NUMBER', 'FundCategory', 'IMPORT_ID']
print(columns)
data = []   #grab results, put into a list, put list into numpy array, and then put numpy array into pandas dataframe
for row in cursor:
    data.append(tuple(row))
df = pd.DataFrame.from_records(np.array(data),columns=columns)
#print(df.shape)

filepath = "C:\\Users\\skirkpatrick\\Coding\\Python\\"
fl = filepath + "RE YTD Cash Receipts thru 6.30.2018.xlsx"
with xlsxwriter.Workbook(fl, {'nan_inf_to_errors': True}) as wb:
    fmt_money = wb.add_format({'num_format': '$#,##0.00'})
    fmt_date = wb.add_format({'num_format': 'mm/dd/yyyy'})
    fmt_dataheader = wb.add_format({'bold': True, 'bg_color': '#000000', 'font_color': '#FFFFFF' })
    fmt_total = wb.add_format({'bg_color': '#434343', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1, 'font_color': '#FFFFFF'})   #darkest grey, white bold text

    #new worksheet - YTD Summary - Appeal Totals
    groupedby_Appeal = df.groupby('Appeal').agg({'SplitAmount':'sum'})
    ws = wb.add_worksheet("YTD Sum - Appeal");
    ws.set_landscape()
    ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
    startingdatarow = 1     #indicates which row to start writing data to
    ws.repeat_rows(startingdatarow-1) #repeats header row on each page for printing (r-1 because it uses excel row numbers, not 0-index rows)
    #writer header
    ws.write(startingdatarow-1,0,'Appeal ID',fmt_dataheader)
    ws.write(startingdatarow-1,1,'RE Total',fmt_dataheader)
    #write detail rows
    r = startingdatarow
    widths = [10,10]
    for i, row in groupedby_Appeal.iterrows():
        r, widths = writeYTDsummary_appeal(ws,r,i,row,widths)
    ws.set_column(0,0,widths[0])
    ws.set_column(1,1,widths[1])

    #new worksheet - YTD Summary, a report that totals payments against checks, grouped by Appeal ID
    groupedby_AppealFund = df.groupby(['Appeal','Fund']).agg({'SplitAmount':'sum'})
    ws = wb.add_worksheet("YTD Sum - Appeal w Fund");
    ws.set_landscape()
    ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
    startingdatarow = 1     #indicates which row to start writing data to
    ws.repeat_rows(startingdatarow-1) #repeats header row on each page for printing (r-1 because it uses excel row numbers, not 0-index rows)
    #writer header
    ws.write(startingdatarow-1,0,'Appeal ID',fmt_dataheader)
    ws.write(startingdatarow-1,1,'Fund',fmt_dataheader)
    ws.write(startingdatarow-1,2,'Total',fmt_dataheader)
    #write detail rows
    r = startingdatarow
    widths = [10,10,10]
    for i, row in groupedby_AppealFund.iterrows():
        r, widths = writeYTDsummary_appealfund(ws,r,i,row,widths)
    ws.set_column(0,0,widths[0])
    ws.set_column(1,1,widths[1])
    ws.set_column(2,2,widths[2])

    #new worksheets - one sheet per Fund Category (Arts Community, Arts Education, Designated, Holding, Right Brain, Undesignated)
    groupedby_FundCat = df.groupby('FundCategory')
    for i, row in groupedby_FundCat:
        ws = wb.add_worksheet("{:s}".format(i))
        ws.set_landscape()
        ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
        startingdatarow = 1     #indicates which row to start writing data to
        ws.repeat_rows(startingdatarow-1) #repeats header row on each page for printing (r-1 because it uses excel row numbers, not 0-index rows)
        #writer header
        ws.write(startingdatarow-1,0,'Appeal ID',fmt_dataheader)
        ws.write(startingdatarow-1,1,'RE',fmt_dataheader)
        ws.write(startingdatarow-1,2,'Albia',fmt_dataheader)
        ws.write(startingdatarow-1,3,'Variance',fmt_dataheader)
        r = startingdatarow   #row counter
#        #initialize length counters for column width
        widths = [10,10]
        groupedby_AppID = row.groupby('Appeal').agg({'SplitAmount':'sum'})
        for j, datarow in groupedby_AppID.iterrows():
            r, widths = writeYTDsummary_appeal(ws,r, j, datarow, widths)
        ws.set_column(0,0,widths[0])
        ws.set_column(1,1,widths[1])

    ws = wb.add_worksheet("Details")
    ws.set_landscape()
    ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
    startingdatarow = 1     #indicates which row to start writing data to
    ws.repeat_rows(startingdatarow-1) #repeats header row on each page for printing (r-1 because it uses excel row numbers, not 0-index rows)
    #writer header
    #['Campaign', 'Fund', 'Appeal', 'GiftID', 'RECORDS_ID', 'Type', 'GiftType', 'TotalAmount', 'SplitAmount', 'AnonRecord', 'AnonGift', 'CONSTITUENT_ID', 'FIRST_NAME', 'KEY_NAME', 'Ref', 'Name', 'Reference', 'POST_DATE', 'CHECK_DATE', 'CHECK_NUMBER', 'GiftDate', 'BATCH_NUMBER', 'FundCategory', 'IMPORT_ID']
    ws.write(startingdatarow-1,0,'Fund Category',fmt_dataheader)
    ws.write(startingdatarow-1,1,'Appeal ID',fmt_dataheader)
    ws.write(startingdatarow-1,2,'Fund',fmt_dataheader)
    ws.write(startingdatarow-1,3,'Name',fmt_dataheader)
    ws.write(startingdatarow-1,4,'Reference',fmt_dataheader)
    ws.write(startingdatarow-1,5,'Amount',fmt_dataheader)
    ws.write(startingdatarow-1,6,'Gift ID',fmt_dataheader)
    ws.write(startingdatarow-1,7,'Gift Date',fmt_dataheader)
    ws.write(startingdatarow-1,8,'Gift Type',fmt_dataheader)
    r = startingdatarow
    widths = [10,10,10,10,10,10,10,15,10]
    for k, row in df.iterrows():
        r, widths  = writedetailrow(ws,r,row,widths)
    ws.set_column(0,0,widths[0])
    ws.set_column(1,1,widths[1])
    ws.set_column(2,2,widths[2])
    ws.set_column(3,3,widths[3])
    ws.set_column(4,4,widths[4])
    ws.set_column(5,5,widths[5])
    ws.set_column(6,6,widths[6])
    ws.set_column(7,7,widths[7])
    ws.set_column(8,8,widths[8])

wb.close()
print("Done!")