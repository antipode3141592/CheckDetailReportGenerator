#   XLSX Report Generator
#   
#   Creates a formatted Excel .xlsx summary report with subtotals from an input file
#   
#   

#   Tested using    - Anaconda 5.0.0
#                   - pandas 0.22.0
#                   - XlsxWriter 1.0.2
#                   - DateTime 4.2
#   IDE: Visual Studio 2017 Community Edition

#import datetime as dt
#import datetime
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

filepath = "C:\\Users\\Antipode\\Documents\\Python Scripting\\"
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
    fl = filepath + str(group1.iloc[0,6]) + " - check " + str(group1.iloc[0,1]) + ".xlsx"
    print(fl)
    #fl = filepath + 'test' + str(name1) + '.xlsx'
    with xlsxwriter.Workbook(fl, {'nan_inf_to_errors': True}) as wb:
        ws = wb.add_worksheet('Report')
        #add formats
        #wb = addformats(wb)
        header1 = "&CCheck Detail Report"
        footer1 = "&L&F" + "&RPage &P of &N"
        ws.set_header(header1)
        ws.set_footer(footer1)
        fmt_money = wb.add_format({'num_format': '$#,##0.00'})
        fmt_date = wb.add_format({'num_format': 'mm/dd/yyyy'})
        fmt_bold = wb.add_format({'bold': True})
        fmt_designated = wb.add_format({'bg_color': '#38761d', 'bold': True, 'num_format': '$#,##0.00'})   #dark green
        fmt_designated_2 = wb.add_format({'bg_color': '#6aa84f', 'bold': True, 'num_format': '$#,##0.00'}) #light green
        fmt_rightbrain = wb.add_format({'bg_color': '#674ea7', 'bold': True, 'num_format': '$#,##0.00'})   #dark purple
        fmt_rightbrain_2 = wb.add_format({'bg_color': '#8e7cc3', 'bold': True, 'num_format': '$#,##0.00'}) #light purple
        fmt_total = wb.add_format({'bg_color': '#ffff00', 'bold': True, 'num_format': '$#,##0.00'}) #yellow

        ws.write(0,0, "Check Detail Report")
        ws.write(3,0, "Check Number: ")
        ws.write(3,1, str(group1.iloc[0,1]))
        ws.write(4,0, "Check Date: ")
        ws.write(4,1, str(group1.iloc[0,2]), fmt_date)
        #writer header
        ws.write(6,0,'Name',fmt_bold)
        ws.write(6,1,'Reference',fmt_bold)
        ws.write(6,2,'Appeal ID',fmt_bold)
        ws.write(6,3,'Date',fmt_bold)
        ws.write(6,4,'Fund',fmt_bold)
        ws.write(6,5,'Gift ID',fmt_bold)
        ws.write(6,6,'Amount',fmt_bold)
        startingdatarow = 7
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
                writesubtotal2(ws,name2,name3,fmt_designated_2,r,firstappealrow,r-1)
                r += 1
            writesubtotal(ws,name2,fmt_designated,r,firstdatarow,r-1)
            r += 1
        writetotal(ws,fmt_total,r,startingdatarow,r-2)
        print("total number of rows:" + str(r))
        #set column widths
        ws.set_column(0,0,columnwidths[0])
        ws.set_column(1,1,columnwidths[1])
        ws.set_column(2,2,columnwidths[2])
        ws.set_column(3,3,columnwidths[3])
        ws.set_column(4,4,columnwidths[4])
        ws.set_column(5,5,columnwidths[5])
        ws.set_column(6,6,columnwidths[6])
        print("closing workbook")
    wb.close()

#ws.set_column('Fund Split Amount', fmt_money)
## Add total rows
#for column in range(6, 11):
#    # Determine where we will place the formula
#    cell_location = xl_rowcol_to_cell(number_rows+1, column)
#    # Get the range to use for the sum formula
#    start_range = xl_rowcol_to_cell(1, column)
#    end_range = xl_rowcol_to_cell(number_rows, column)
#    # Construct and write the formula
#    formula = "=SUM({:s}:{:s})".format(start_range, end_range)
#    worksheet.write_formula(cell_location, formula, total_fmt)
## Add a total label
#worksheet.write_string(number_rows+1, 5, "Total",total_fmt)
#percent_formula = "=1+(K{0}-G{0})/G{0}".format(number_rows+2)
#worksheet.write_formula(number_rows+1, 11, percent_formula, total_percent_fmt)
## Define our range for the color formatting
#color_range = "L2:L{}".format(number_rows+1)
## Add a format. Light red fill with dark red text.
#format1 = workbook.add_format({'bg_color': '#FFC7CE',
#                               'font_color': '#9C0006'})
## Add a format. Green fill with dark green text.
#format2 = workbook.add_format({'bg_color': '#C6EFCE',
#                               'font_color': '#006100'})

#df1a = df1[df1['Gift Batch Number'].values == 6822]