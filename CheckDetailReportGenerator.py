#   XLSX Report Generator
#   Copyright 2018 by Sean Vo Kirkpatrick using GNU GPL v3
#   skirkpatrick@racc.org or seanvokirkpatrick@gmail.com
#   
#   Creates a formatted Excel .xlsx summary report with subtotals, data pulled from Raiser's Edge DB directly
#       using parameterized stored procedure
#   Input:  a date variable
#   Output: a bunch of reports
#
#   Tested using    - Anaconda 5.2.0
#                   - pandas 0.23.0
#                   - XlsxWriter 1.0.4
#                   - pyodbc 4.0.23
#                   - numpy 1.15.2
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

#-------------------------------------------------------------------------------
# Input:  Select the date range (for Gift Date) that you wish to create reports for
startdate = '2018-12-28'         
enddate = '2018-12-28'         
filepath = "C:\\Users\\skirkpatrick\\Coding\\Python\\"  #output path for the generated files
#-------------------------------------------------------------------------------

#------------------------------------------------------------------------------------------------------
# define those functions!
def writedetailrow(ws,r,row,width):
    #didn't use a loop for the comparisons because some columns don't need resizing
    if len(str(row[15])) > width[0]:
        width[0] = len(str(row[15]))
    if len(str(row[14])) > width[1]:
        width[1] = len(str(row[14]))
    if len(str(row[3])) > width[2]:
        width[2] = len(str(row[3]))
    if len(str(row[2])) > width[4]:
        width[4] = len(str(row[2]))
    if len(str(row[4])) > width[5]:
        width[5] = len(str(row[4]))
    ws.write(r,0,row[15])
    ws.write(r,1,row[14])
    ws.write(r,2,row[3])
    ws.write(r,3,str(row[19]))
    ws.write(r,4,row[2])
    ws.write(r,5,row[4])
    ws.write(r,6,row[8],fmt_money)
    return width

def writesubtotal(ws,groupname, fmt,r,row_1st,row_last):
    ws.write(r,0,"Subtotal - " + str(groupname), fmt)
    ws.write(r,1,"", fmt)
    ws.write(r,2,"", fmt)
    ws.write(r,3,"", fmt)
    ws.write(r,4,"", fmt)
    ws.write(r,5,"", fmt)
    formula = "=SUBTOTAL(109,{:s})".format(xl_range(row_1st,6,row_last,6))
    ws.write_formula(xl_rowcol_to_cell(r,6), formula, fmt)
    return

def writesubtotal2(ws,groupname1, groupname2, fmt,r,row_1st,row_last):
    ws.write(r,0,"Subtotal - " + str(groupname1) + " - " + str(groupname2), fmt)
    ws.write(r,1,"", fmt)
    ws.write(r,2,"", fmt)
    ws.write(r,3,"", fmt)
    ws.write(r,4,"", fmt)
    ws.write(r,5,"", fmt)
    formula = "=SUBTOTAL(109,{:s})".format(xl_range(row_1st,6,row_last,6))
    ws.write_formula(xl_rowcol_to_cell(r,6), formula, fmt)
    return

def writetotal(ws,fmt,r,row_1st,row_last):
    ws.write(r,0,"Total",fmt)
    ws.write(r,1,"", fmt)
    ws.write(r,2,"", fmt)
    ws.write(r,3,"", fmt)
    ws.write(r,4,"", fmt)
    ws.write(r,5,"", fmt)
    formula = "=SUBTOTAL(109,{:s})".format(xl_range(row_1st,6,row_last,6))
    ws.write_formula(xl_rowcol_to_cell(r,6), formula, fmt)
    return

# replaces restricted system characters with underscores
def sterilizestring (s):
    s = str(s)
    for char in "?.!/;:":
        s = s.replace(char,'_');
    return s
#------------------------------------------------------------------------------------------------------

#connect to db, requires windows integrated security for this connection string
cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};" #requires explicitily stating the sql driver
                      "Server=overlook;"
                      "Database=re_racc;"
                      "Trusted_Connection=yes;")    #use windows integrated security
cursor = cnxn.cursor()

sqlcommand = 'exec sp_checkdetailreport ''?'', ''?'' '  #call stored procedure
#note that the FundCategory sorting is done in the stored procudure.  the order is:
# ["Designated", "Arts Ed", "Right Brain", "RACC", "Community", "Holding", "Undesignated"]
sqlparams = (startdate,enddate)
cursor.execute(sqlcommand,sqlparams)
columns = [column[0] for column in cursor.description]

data = []   #grab results, put into a list, put list into numpy array, and then put numpy array into pandas dataframe
for row in cursor:
    data.append(tuple(row))
    #print(row)
df = pd.DataFrame.from_records(np.array(data),columns=columns)
print("Query results count(rows): " + str(df.shape[0]))

groupedby_Batch = df.groupby('BATCH_NUMBER')
for name1, group1 in groupedby_Batch:
    #if group1.iloc[0,1] != "":
    print("---------------------")
    print("Batch: " + str(name1))
    print(group1)
    if (str(group1.iloc[0,17]) != "" and str(group1.iloc[0,17]) != "None") and (str(group1.iloc[0,2]) != ""):
        print("Appeal ID: " + group1.iloc[0,2] + ", Check# " + str(group1.iloc[0,19]))
        fl = filepath + "{:s}".format(sterilizestring(group1.iloc[0,2])) + " - check {:s}".format(str(group1.iloc[0,17])) + ".xlsx"
    elif (str(group1.iloc[0,17]) == "None"):
        fl = filepath + "{:s}".format(sterilizestring(group1.iloc[0,2])) + " - Batch {:s}".format(name1) + ".xlsx"
    else:
        print("no check number!  using truncated filename")
        fl = filepath + "{:s}".format(sterilizestring(group1.iloc[0,2])) + ".xlsx"
    print("Filename: " + fl)
    summary = group1.groupby(['Fund','Appeal']).sum()
    with xlsxwriter.Workbook(fl, {'nan_inf_to_errors': True}) as wb:
        ws = wb.add_worksheet('Report')
        #add formats
        header1 = "&CCheck Detail Report"
        footer1 = "&L&Z&F" + "&RPage &P of &N"
        ws.set_header(header1)
        ws.set_footer(footer1)
        ws.set_landscape()
        ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
        fmt_money = wb.add_format({'num_format': '$#,##0.00'})
        fmt_date = wb.add_format({'num_format': 'mm/dd/yyyy'})
        fmt_dataheader = wb.add_format({'bold': True, 'bg_color': '#000000', 'font_color': '#FFFFFF' })
        fmt_subtotal = wb.add_format({'bg_color': '#b7b7b7', 'bold': True , 'num_format': '$#,##0.00', 'bottom':1, 'top':1}) #light grey
        fmt_subtotal2 = wb.add_format({'bg_color': '#666666', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1})   #dark grey
        fmt_total = wb.add_format({'bg_color': '#434343', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1, 'font_color': '#FFFFFF'})   #darkest grey, white bold text

        ws.write(0,0, "Check Number: ")
        if group1.iloc[0,19] != "":
            ws.write(0,1, '{:s}'.format(str(group1.iloc[0,17])))
        else:
            ws.write(0,1, " --- ")
        ws.write(1,0, "Check Date: ")
        if group1.iloc[0,18] != "":
            ws.write(1,1, '{:s}'.format(str(group1.iloc[0,16])))
        else:
            ws.write(1,1, " --- ")
        #Rollup report for top of report
        #print list of categories and appeals
        r=1
        ws.write(0,3,'Category',fmt_dataheader)
        ws.write(0,4,'Appeal ID',fmt_dataheader)
        ws.write(0,5,'Subtotals',fmt_dataheader)
        ws.write(0,6,'Totals',fmt_dataheader)
        #summed = group1.groupby(['Fund Category', 'Appeal ID'])
        group_category = group1.groupby('FundCategory', sort=False)
        for n1,g1 in group_category:
            ws.write(r,3,n1)
            group_appeal = g1.groupby('Appeal')
            for n2,g2 in group_appeal:
                ws.write(r,4,n2)
                r += 1
            ws.write(r,4,n1,fmt_subtotal)
            ws.write(r,5,"",fmt_subtotal)
            r += 1
        startingdatarow = r + 3     #indicates which row to start writing data to
        if(startingdatarow < 4):
            startingdatarow = 4
        ws.repeat_rows(startingdatarow-1) #repeats header row on each page for printing (r-1 because it uses excel row numbers, not 0-index rows)
        #writer header
        ws.write(startingdatarow-1,0,'Name',fmt_dataheader)
        ws.write(startingdatarow-1,1,'Reference',fmt_dataheader)
        ws.write(startingdatarow-1,2,'Appeal ID',fmt_dataheader)
        ws.write(startingdatarow-1,3,'Date',fmt_dataheader)
        ws.write(startingdatarow-1,4,'Fund',fmt_dataheader)
        ws.write(startingdatarow-1,5,'Gift ID',fmt_dataheader)
        ws.write(startingdatarow-1,6,'Amount',fmt_dataheader)
        r = startingdatarow   #row counter
        row_subtotals = 1
        subgroup = group1.groupby('FundCategory',sort=False)
        #initialize length counters for column width
        column_widths = [19,10,10,11,20,8,10]
        for name2, group2 in subgroup:
            print("Current Group: " + str(name2) + " on row " + str(r))
            firstdatarow = r    #preserve the 1st row number of each Fund Category group
            appealgroup = group2.groupby('Appeal')
            for name3, group3 in appealgroup:
                firstappealrow = r      #preserve the 1st row number of each Appeal group
                print("Current Appeal: " + str(name3) + " on row " + str(r))
                for row in group3.itertuples():
                    column_widths = writedetailrow(ws,r,row,column_widths)
                    r += 1
                writesubtotal2(ws,name2,name3,fmt_subtotal,r,firstappealrow,r-1)
                formula = "={:s}".format(xl_rowcol_to_cell(r,6))
                ws.write_formula(xl_rowcol_to_cell(row_subtotals,5),formula,fmt_money)
                row_subtotals += 1
                r += 1
            writesubtotal(ws,name2,fmt_subtotal2,r,firstdatarow,r-1)
            formula = "={:s}".format(xl_rowcol_to_cell(r,6))
            ws.write_formula(xl_rowcol_to_cell(row_subtotals,6),formula,fmt_subtotal)
            row_subtotals += 1
            r += 1
        writetotal(ws,fmt_total,r,startingdatarow,r-2)
        formula = "={:s}".format(xl_rowcol_to_cell(r,6))
        ws.write(row_subtotals,5,"Total",fmt_total)
        ws.write_formula(xl_rowcol_to_cell(row_subtotals,6),formula,fmt_total)
        r+=3
        ws.write(r,0,"Reported by Sean K. on {}".format(dt.datetime.now().strftime("%m/%d/%y")))
        ws.set_column(0,0,column_widths[0])
        ws.set_column(1,1,column_widths[1])
        ws.set_column(2,2,column_widths[2])
        ws.set_column(3,3,column_widths[3])
        ws.set_column(4,4,column_widths[4])
        ws.set_column(5,5,column_widths[5])
        ws.set_column(6,6,column_widths[6])
        print("closing workbook")
    wb.close()