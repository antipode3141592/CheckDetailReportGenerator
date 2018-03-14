#   XLSX Report Generator
#   Copyright 2018 by Sean Vo Kirkpatrick using GNU GPL v3
#   skirkpatrick@racc.org or sean@studioantipode.com or seanvokirkpatrick@gmail.com
#   
#   Creates a formatted Excel .xlsx summary report with subtotals from each group of Appeals in an input file
#
#   Tested using    - Anaconda 5.0.0
#                   - pandas 0.22.0
#                   - XlsxWriter 1.0.2
#                   - xlrd 1.1.0
#                   - openpyxl 2.5.1
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
import os
import openpyxl
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
#from openpyxl import Worksheet
#import xlsxwriter
#from xlsxwriter.utility import xl_rowcol_to_cell
#from xlsxwriter.utility import xl_range

def write_detailrow(r,row,ws):
    ws.cell(row=r,column=1,value="{0}".format(row[11]))
    ws.cell(row=r,column=2,value="{0}".format(row[12]))
    ws.cell(row=r,column=3,value="{0}".format(row[13]))
    ws.cell(row=r,column=4,value="{0}".format(row[14]))
    ws.cell(row=r,column=5,value="{0}".format(row[15]))
    ws.cell(row=r,column=6,value="{0}".format(row[16]))
    a = ws.cell(row=r,column=7,value="{0}".format(row[17]))
    a.number_format = '$#,##0.00'
    #c.style(Font(bold=True))
    #c.number_style("$#,#00.00")
    r+=1    
    return(r)

def write_summaryrow(r,_r,row,ws):
    ws.cell(row=r,column=1,value="Subtotal - {0}".format(row[8]))
    ws.cell(row=r,column=2)
    ws.cell(row=r,column=3)
    ws.cell(row=r,column=4)
    ws.cell(row=r,column=5)
    ws.cell(row=r,column=6)
    ws.cell(row=r,column=7,value="=SUBTOTAL(109,G{0}:G{1})".format(_r,r-1))
    r+=1
    return(r)

#print("oh hai")
filepath = "C:\\Users\\Antipode\\Documents\\Python Scripting\\"
outputpath = "C:\\Users\\Antipode\\Documents\\Python Scripting\\17-18\\"

#filepath = "C:\\Users\\skirkpatrick\\Coding\\Python\\"
#outputpath = "C:\\Users\\skirkpatrick\\Coding\\Python\\17-18\\"
inputfile = "PLEDGE_Q.XLSX"
if os.path.exists(filepath + inputfile):
    print(filepath + inputfile)
else:
    print("no input file!")
xl = pd.ExcelFile(filepath + inputfile) #use pandas' excel reader
#   columns of import sheet
#   ['Gift Type',	'Constituent Specific Attributes Arts Card Date',	'Constituent Specific Attributes Arts Card - OLD Date',	'Gift Pledge Balance',	'Gift Is Anonymous',
#   	'Matching Gift Import ID',	'Gift Import ID',	'Fund Category',	'Fund ID',	'Constituent ID',	'Name',	'Appeal ID',	'Gift Reference',	'Gift Date',
#   	'Fund Description',	'Gift ID',	'Fund Split Amount',	'Preferred Address Lines',	'Preferred Address Line 1',	'Preferred Address Line 2',	'Preferred Address Line 3',
#   	'Preferred City_ State',	'Preferred ZIP']

df = xl.parse()
df = df.replace(pd.np.nan, '', regex=True)  #replaces all types of NAN entries with blank space
xl.close()
groupedby_Appeal = df.groupby('Appeal ID')
for name1, group1 in groupedby_Appeal:
    #try openin file
    fp = outputpath + name1 + ".xlsx"
    fp2 = outputpath + "Test\\" + name1 + "_test.xlsx"
    if os.path.exists(fp):
        print("{0} exists!".format(name1))
        wb = openpyxl.load_workbook(fp)
        ws = wb.create_sheet(title="{0}".format(dt.datetime.now().strftime("%m-%d-%y")))
        #Rollup report for top of report
        #print list of categories
        r=2 #openpyxl uses 1-based index, same as excel
        ws.cell(row=1,column=3,value="Category")
        ws.cell(row=1,column=4,value="Subtotals")
        ws.cell(row=1,column=5,value="Totals")
        groupby_category = group1.groupby('Fund Category')
        for n1,g1 in groupby_category:
            ws.cell(column=3,row=r,value=n1)
            r += 1
        startingdatarow = r + 3     #indicates which row to start writing data to
        if(startingdatarow < 4):
            startingdatarow = 4
        r = startingdatarow
        #writer header
        ws.cell(row=startingdatarow-1,column=1,value="Name",)
        ws.cell(row=startingdatarow-1,column=2,value="Reference")
        ws.cell(row=startingdatarow-1,column=3,value="Appeal ID")
        ws.cell(row=startingdatarow-1,column=4,value="Date")
        ws.cell(row=startingdatarow-1,column=5,value="Fund")
        ws.cell(row=startingdatarow-1,column=6,value="Gift ID")
        ws.cell(row=startingdatarow-1,column=7,value="Amount")

        for n1,g1 in groupby_category:
            startingcategoryrow = r
            for row in g1.itertuples():
                r = write_detailrow(r,row,ws)
            r = write_summaryrow(r,startingcategoryrow,row,ws)
            #ws.cell(column=7,row=r,value="=SUBTOTAL(109,G{0}:G{1})".format(startingcategoryrow,r-1))
            #r+=1
        wb.save(fp2)
    else:
        print("{0} not found!".format(fp))
