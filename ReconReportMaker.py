#   XLSX Report Generator
#   Copyright 2018 by Sean Vo Kirkpatrick using GNU GPL v3
#   skirkpatrick@racc.org or sean@studioantipode.com or seanvokirkpatrick@gmail.com
#   
#   Creates a formatted Excel .xlsx summary report with subtotals from each batch in an input file
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
    if len(str(row[4])) > width[0]:
        width[0] = len(str(row[4]))
    if len(str(row[6])) > width[1]:
        width[1] = len(str(row[6]))
    if len(str(row[8])) > width[2]:
        width[2] = len(str(row[8]))
    if len(str(row[12])) > width[3]:
        width[3] = len(str(row[12]))
    if len(str(row[13])) > width[4]:
        width[4] = len(str(row[13]))
    if len(str(row[14])) > width[5]:
        width[5] = len(str(row[14]))
    ws.write(r,0,row[4])    #fund description
    ws.write(r,1,row[6])     #appeal id
    ws.write(r,2,row[8],fmt_date)   #gift date
    ws.write(r,3,row[12])    #name
    ws.write(r,4,row[13])    #reference
    ws.write(r,5,row[14],fmt_money) #amount
    return width

def writetotal(ws,fmt,r,row_1st,row_last):
    ws.write(r,0,"Total",fmt)
    ws.write(r,1,"", fmt)
    ws.write(r,2,"", fmt)
    ws.write(r,3,"", fmt)
    ws.write(r,4,"", fmt)
    ws.write(r,5,"", fmt)
    formula = "=SUBTOTAL(109,{:s})".format(xl_range(row_1st,5,row_last,5))
    ws.write_formula(xl_rowcol_to_cell(r,5), formula, fmt)
    return

def sterilizestring (s):
    for char in "?.!/;:":
        s = s.replace(char,'_');
    return s

#filepath = "C:\\Users\\Antipode\\Documents\\Python Scripting\\"
filepath = "C:\\Users\\skirkpatrick\\Coding\\Python\\"
outputpath = "C:\\Users\\skirkpatrick\\Coding\\Python\\Outgoing\\"
inputfile = "CHECK.XLSX"
xl = pd.ExcelFile(filepath + inputfile) #use pandas' excel reader
#   columns of import sheet
#   [Anonymous?,	Gift Is Anonymous,	Gift Type,	Fund Description,	Fund Address,	Appeal ID,	Gift ID,	Gift Date,	Constituent ID,	_name,	
#       Gift Reference,	Name,	Reference,  Fund Split Amount]

df = xl.parse()
df = df.replace(pd.np.nan, '', regex=True)  #replaces all types of NAN entries with blank space
xl.close()
groupedby_Fund = df.groupby('Fund Description')
for name1, group1 in groupedby_Fund:
    localsum = group1['Fund Split Amount'].sum()
    if localsum < 20.00:
        print("Skipping group " + name1 + ", with sum of $" + str(localsum))
        continue;
    fl = outputpath + "{:s}".format(sterilizestring(group1.iloc[0,3])) + ".xlsx"
    print("Group: " + name1 + " , sum $" + str(localsum))
    with xlsxwriter.Workbook(fl, {'nan_inf_to_errors': True}) as wb:
        ws = wb.add_worksheet('Report')
        #add formats, header, and footer
        header1 = "&CWork for Art - Payments Received on Pledges"
        footer1 = "&RPage &P of &N"
        ws.set_header(header1)
        ws.set_footer(footer1)
        ws.set_landscape()
        ws.fit_to_pages(1,0)    #printing is 1 page wide, no limit on height/length
        fmt_money = wb.add_format({'num_format': '$#,##0.00'})
        fmt_date = wb.add_format({'num_format': 'mm/dd/yyyy'})
        fmt_dataheader = wb.add_format({'bold': True, 'bg_color': '#000000', 'font_color': '#FFFFFF' })
        fmt_total = wb.add_format({'bg_color': '#434343', 'bold': True, 'num_format': '$#,##0.00', 'bottom':1, 'top':1, 'font_color': '#FFFFFF'})   #darkest grey, white bold text

        startingdatarow = 1     #indicates which row to start writing data to
        ws.repeat_rows(startingdatarow-1) #repeats header row on each page for printing (r-1 because it uses excel row numbers, not 0-index rows)
        #writer header
        ws.write(startingdatarow-1,0,'Fund',fmt_dataheader)
        ws.write(startingdatarow-1,1,'Appeal ID',fmt_dataheader)
        ws.write(startingdatarow-1,2,'Gift Date',fmt_dataheader)
        ws.write(startingdatarow-1,3,'Name',fmt_dataheader)
        ws.write(startingdatarow-1,4,'Reference',fmt_dataheader)
        ws.write(startingdatarow-1,5,'Amount',fmt_dataheader)
        r = startingdatarow   #row counter
        row_subtotals = 1
        #initialize length counters for column width
        column_widths = [19,10,11,20,8,10]
        for row in group1.itertuples():
            column_widths = writedetailrow(ws,r,row,column_widths)
            r += 1
        writetotal(ws,fmt_total,r,startingdatarow,r-1)
        r+=3
        ws.write(r,0,"Reported by Sean Kirkpatrick on {}".format(dt.datetime.now().strftime("%m/%d/%y")))
        ws.set_column(0,0,column_widths[0])
        ws.set_column(1,1,column_widths[1])
        ws.set_column(2,2,column_widths[2])
        ws.set_column(3,3,column_widths[3])
        ws.set_column(4,4,column_widths[4])
        ws.set_column(5,5,column_widths[5])
        #print("closing workbook")
    wb.close()
