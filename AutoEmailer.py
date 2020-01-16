#   AutoEmailer
#   Copyright 2018 by Sean Vo Kirkpatrick using GNU GPL v3
#   skirkpatrick@racc.org or sean@studioantipode.com or seanvokirkpatrick@gmail.com
#
#   Sends an e-mail with an attachment, based on the file name and a lookup table in an excel file,
#   for each attachment file in a given input folder.
#   
#   Requirements:     
#               -a running and e-mail ready Outlook instance
#               -Excel file of contacts to e-mail
#               -Files to attach to e-mails in a folder
#               -e-mail body text in a text file
#   Outputs:    -Lists e-mails sent
#               -Lists contacts not found
#
#   Files to be attached must have names of the form "[Org Name] - [some text].[some filetype]"
#       Program doesn't care what filetypes are used as attachments, but all of the rules Outlook enforces
#       will be enforced, such as file size and type.
#   Makes a connection to an open instance of Outlook, e-mails will be sent using the active account
#   The program loops through all of the files in the "path_input" folder:
#       First, parse the file name name to get the organization name.
#       Second, search the contact list excel file for the appropriate e-mail addresses.
#       Send the file as an attachment to all e-mail addresses found, with a body set to the "file_emailbody" file
#       (note: currently only plain text)
#   Files that have a positive number of associated e-mail addresses are then moved to the "path_output" folder, 
#       and files that had no associated e-mail addresses to the "path_errors" folder.

#
#   Tested using    - Anaconda 5.0.0
#                   - pandas 0.22.0
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

import win32com.client as win32
import pyodbc 
import os
import numpy as np
import pandas as pd

#small function for replacing os-restricted characters in filenames
def sterilizestring (s):
    for char in "?.!/;:":
        s = s.replace(char,'_');
    return s


outlook = win32.DispatchEx('outlook.application')   #TODO: modify to use any e-mail, not just control Outlook

#file paths
path_root = "C:\\Users\\skirkpatrick\\Documents\\Work for Art\\DG Pledge Notification Reports\\"
#path_contactlist = "//concordia/lancentral/Work for Art/Designated Gifts/DG Pledge Reports//"   #network locations use forward slashes
path_input = "C:\\Users\\skirkpatrick\\Documents\\Work for Art\\DG Pledge Notification Reports\\Input\\"
path_output = "C:\\Users\\skirkpatrick\\Documents\\Work for Art\\DG Pledge Notification Reports\\Output\\"
path_errors = "C:\\Users\\skirkpatrick\\Documents\\Work for Art\\DG Pledge Notification Reports\\Errors\\"
file_emailbody = path_root + "email template.txt"

##this contact list is populated by 
#file_contactlist = path_root + "DG Emails.xlsx"
#populate contact list dataframe from database
cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};" #requires explicitily stating the sql driver
                      "Server=overlook;"
                      "Database=re_racc;"
                      "Trusted_Connection=yes;")    #use windows integrated security
cursor = cnxn.cursor()
#note: turned the query into a stored procedure for better security (but there aren't parameters, so probably unnecessary?)
#cursor.execute("select a.DESCRIPTION, a.FUND_ID, b.LONGDESCRIPTION as Category, d.CONSTITUENT_ID, d.ORG_NAME, e.num as Email, f.LONGDESCRIPTION as Type from FUND a join TABLEENTRIES b on (a.FUND_CATEGORY = b.TABLEENTRIESID) join FUND_ORG_RELATIONSHIPS c on (a.ID = c.FUND_ID) join RECORDS d on (c.CONSTIT_ID = d.ID) join PHONES e on (e.CONSTIT_ID = d.ID) join TABLEENTRIES f on (e.PHONETYPEID =  f.TABLEENTRIESID) where b.LONGDESCRIPTION = 'Designated' and f.LONGDESCRIPTION like 'E-Mail%' and e.CONSTIT_RELATIONSHIPS_ID is null order by a.DESCRIPTION, f.LONGDESCRIPTION")
cursor.execute("sp_getdgfundemails")
data = []   #grab results, put into a list, put list into numpy array, and then put numpy array into pandas dataframe
for row in cursor:
    data.append(tuple(row))
contactlist = pd.DataFrame.from_records(np.array(data))

#load data
message_body = open(file_emailbody).read()  #load message body from file
#wb = pd.ExcelFile(file_contactlist)
#df1 = wb.parse('Contacts')  #parse contents of excel file into Pandas dataframe

#initialize lists
files_good = []
files_error = []

#email loop
for f in os.listdir(path_input):
    _f = f.split(" - ")                 #split filename of form "Org Name - New Pledges through Work for Art.xlxs" into array _f
    validemailacount = 0
    emaillist = []      #list of e-mails to send to (each item might need to be mailed to multiple addresses)
    #for index, row in df1.iterrows():   #iterate over all rows, searching for matches
    #    if row[0] == _f[0]:             #if column 1 (org abbreviation) matches 1st part of File Name, add e-mail to list
    #        emaillist.append(row[2])    #linear search, iterate overy every row, slow, but workable
    #        validemailacount += 1
    for index, row in contactlist.iterrows():   #iterate over all rows, searching for matches
        if sterilizestring(row[1]) == sterilizestring(_f[0]):             #if Org Name matches File Name, add e-mail to list
            emaillist.append(row[5])
            ach = row[7]
            validemailacount += 1
    if validemailacount == 0:    #if no valid matches, print error and save filename in error list
        files_error.append(f)
        print(_f[0], " not found!")
    else:   #if there is at least one matching e-mail, send e-mail
        stringemail = ""
        print("Org: ", _f[0])
        files_good.append(f)    #add filename to good list
        for address in emaillist:
            stringemail += address + ";"
        print("Sending email to: ", stringemail)
        mail = outlook.CreateItem(0)
        mail.To = stringemail
        mail.Subject = 'Pledges Received through Arts Impact Fund'
        mail.Body = message_body
        mail.Attachments.Add(path_input + f) #f is the excel report filename
        mail.Send()

#file moving loops
for f in files_good:
    os.rename(path_input + f, path_output + f)
for f in files_error:
    os.rename(path_input + f, path_errors + f)