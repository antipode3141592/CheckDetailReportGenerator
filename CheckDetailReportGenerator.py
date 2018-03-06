#   XLSX Report Generator
#   Copyright 2018 by Sean Vo Kirkpatrick using GNU GPL v3
#   skirkpatrick@racc.org or sean@studioantipode.com or seanvokirkpatrick@gmail.com
#   
#   Creates a formatted Excel .xlsx summary report with subtotals from an input file
#   
#   

#   Tested using    - Anaconda 5.0.0
#                   - pandas 0.22.0
#                   - XlsxWriter 1.0.2
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

# define those functions!
def writedetailrow(ws,r,row,width):
    #didn't use a loop for the comparisons because some columns don't need resizing
    if len(str(row[5])) > width[0]:
        width[0] = len(str(row[5]))
    if len(str(row[6])) > width[1]:
        width[1] = len(str(row[6]))
    if len(str(row[7])) > width[2]:
        width[2] = len(str(row[7]))
    if len(str(row[9])) > width[4]:
        width[4] = len(str(row[9]))
    if len(str(row[10])) > width[5]:
        width[5] = len(str(row[10]))
    ws.write(r,0,row[5])
    ws.write(r,1,row[6])
    ws.write(r,2,row[7])
    ws.write(r,3,row[8],fmt_date)
    ws.write(r,4,row[9])
    ws.write(r,5,row[10])
    ws.write(r,6,row[11],fmt_money)
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

#filepath = "C:\\Users\\Antipode\\Documents\\Python Scripting\\"
filepath = "C:\\Users\\skirkpatrick\\Coding\\Python\\"
inputfile = "CHECK_DE.XLSX"
xl = pd.ExcelFile(filepath + inputfile) #use pandas' excel reader
#   columns of import sheet
#   ['Gift Batch Number', 'Gift Check Number', 'Check Date', 'Fund Category', 'Name', 
#       'Gift Reference', 'Appeal ID', 'Gift Date', 'Fund Description', 'Gift ID', 'Fund Split Amount']
df = xl.parse()
df = df.replace(pd.np.nan, '', regex=True)  #replaces all types of NAN entries with blank space
xl.close()
groupedby_Batch = df.groupby('Gift Batch Number')
for name1, group1 in groupedby_Batch:
    if group1.iloc[0,1] != "":
        fl = filepath + "{:s}".format(group1.iloc[0,6]) + " - check {:.0f}".format(group1.iloc[0,1]) + ".xlsx"
    else:
        print("no check number!  using truncated filename")
        fl = filepath + "{:s}".format(group1.iloc[0,6]) + ".xlsx"
    print("Filename: " + fl)
    summary = group1.groupby(['Fund Description','Appeal ID']).sum()
    with xlsxwriter.Workbook(fl, {'nan_inf_to_errors': True}) as wb:
        ws = wb.add_worksheet('Report')
        #add formats
        header1 = "&CCheck Detail Report"
        footer1 = "&L&F" + "&RPage &P of &N"
        ws.set_header(header1)
        ws.set_footer(footer1)
        ws.set_landscape()
        ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
        #ws.hide_gridlines(0)    #do not show gridlines
        startingdatarow = 4     #indicates which row to start writing data to
        ws.repeat_rows(startingdatarow) #repeats header row on each page for printing
        fmt_money = wb.add_format({'num_format': '$#,##0.00'})
        fmt_date = wb.add_format({'num_format': 'mm/dd/yyyy'})
        fmt_dataheader = wb.add_format({'bold': True, 'bg_color': '#000000', 'font_color': '#FFFFFF', })
        fmt_artsed = wb.add_format({'bg_color': '#92d050', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1})   #dark green
        fmt_artsed_2 = wb.add_format({'bg_color': '#d8e4bc', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1}) #light green
        fmt_rightbrain = wb.add_format({'bg_color': '#60497a', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1})   #dark purple
        fmt_rightbrain_2 = wb.add_format({'bg_color': '#b1a0c7', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1}) #light purple
        fmt_community = wb.add_format({'bg_color': '#ffff00', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1})   #yellow
        fmt_community_2 = wb.add_format({'bg_color': '#eeedb3', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1}) #light yellow
        fmt_designated = wb.add_format({'bg_color': '#ff9900', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1})   #orange
        fmt_designated_2 = wb.add_format({'bg_color': '#f6b26b', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1}) #light orange
        fmt_raccu = wb.add_format({'bg_color': '#0b5394', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1})   #blue
        fmt_raccu_2 = wb.add_format({'bg_color': '#6fa8dc', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1}) #light blue
        fmt_holding = wb.add_format({'bg_color': '#666666', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1})   #dark grey
        fmt_holding_2 = wb.add_format({'bg_color': '#b7b7b7', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1}) #light grey
        fmt_total = wb.add_format({'bg_color': '#ffc000', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':6}) #light orange

        ws.write(0,0, "Check Number: ")
        if group1.iloc[0,1] != "":
            ws.write(0,1, '{:.0f}'.format(group1.iloc[0,1]))
        else:
            ws.write(0,1, " --- ")
        ws.write(1,0, "Check Date: ")
        if group1.iloc[0,2] != "":
            ws.write(1,1, '{:s}'.format(group1.iloc[0,2]), fmt_date)
        else:
            ws.write(1,1, " --- ")
        #writer header
        ws.write(startingdatarow-1,0,'Name',fmt_dataheader)
        ws.write(startingdatarow-1,1,'Reference',fmt_dataheader)
        ws.write(startingdatarow-1,2,'Appeal ID',fmt_dataheader)
        ws.write(startingdatarow-1,3,'Date',fmt_dataheader)
        ws.write(startingdatarow-1,4,'Fund',fmt_dataheader)
        ws.write(startingdatarow-1,5,'Gift ID',fmt_dataheader)
        ws.write(startingdatarow-1,6,'Amount',fmt_dataheader)
        r = startingdatarow   #row counter
        subgroup = group1.groupby('Fund Category')
        #initialize length counters for column width
        columnwidths = [19,10,10,11,20,8,10]
        for name2, group2 in subgroup:
            print("Current Group: " + str(name2) + " on row " + str(r))
            firstdatarow = r    #preserve the 1st row number of each Fund Category group
            appealgroup = group2.groupby('Appeal ID')
            for name3, group3 in appealgroup:
                firstappealrow = r      #preserve the 1st row number of each Appeal group
                print("Current Appeal: " + str(name3) + " on row " + str(r))
                for row in group3.itertuples():
                    columnwidths = writedetailrow(ws,r,row,columnwidths)
                    r += 1
                if name2 == "Designated":
                    _fmt2 = fmt_designated_2
                elif name2 == "Community":
                    _fmt2 = fmt_community_2
                elif name2 == "Arts Ed":
                    _fmt2 = fmt_artsed_2
                elif name2 == "Right Brain":
                    _fmt2 = fmt_rightbrain_2
                elif name2 == "Holding":
                    _fmt2 = fmt_holding_2
                elif name2 == "RACC":
                    _fmt2 = fmt_raccu_2
                elif name2 == "Holding":
                    _fmt2 = fmt_holding_2
                writesubtotal2(ws,name2,name3,_fmt2,r,firstappealrow,r-1)
                r += 1
            if name2 == "Designated":
                _fmt = fmt_designated
            elif name2 == "Community":
                _fmt = fmt_community
            elif name2 == "Arts Ed":
                _fmt = fmt_artsed
            elif name2 == "Right Brain":
                _fmt = fmt_rightbrain
            elif name2 == "Holding":
                _fmt = fmt_holding
            elif name2 == "RACC":
                _fmt = fmt_raccu
            elif name2 == "Holding":
                _fmt = fmt_holding
            writesubtotal(ws,name2,_fmt,r,firstdatarow,r-1)
            r += 1
        writetotal(ws,fmt_total,r,startingdatarow,r-2)
        r+=3
        ws.write(r,0,"Reported by Sean K. on {}".format(dt.datetime.now().strftime("%m/%d/%y")))
        ws.set_column(0,0,columnwidths[0])
        ws.set_column(1,1,columnwidths[1])
        ws.set_column(2,2,columnwidths[2])
        ws.set_column(3,3,columnwidths[3])
        ws.set_column(4,4,columnwidths[4])
        ws.set_column(5,5,columnwidths[5])
        ws.set_column(6,6,columnwidths[6])
        print("closing workbook")
    wb.close()