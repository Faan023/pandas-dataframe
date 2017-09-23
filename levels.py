import MySQLdb
import pandas as pd
from openpyxl import load_workbook
import os

'''
In this business there are several sales level challenges.
These challenges can run over one week, several weeks or even months.
Let's give an example of a challenge:
Starting in week 30, ending in week 32:
From 3500 sales, every 1500 after 3500, you get ticket into a draw for a
prize at the end of the challenge. Instead of having to go through thousands
of invoices and customer accounts,(usually done by two or three people)to determine
how many draw tickets every one of the 7000 strong sales force has to get,
I came up with this.
It simply makes a summary of the total sales of each person in the specified
time, and gives the amount of tickets that needs to be allocated to that person,
saving hours of painstaking work so more challenges can be run, motivating
the salesforce and increasing turnover.
'''


#create a connection to MySQL

mysql_cn = MySQLdb.connect(host="hostname",
                           port=3306,user="root",passwd="password",
                           db="my_db_name")
# connect to MySQL
df = pd.read_sql("Select Week_No, SF_Code, SF_Name, Recognition from invoiceregister;",con=mysql_cn)

# Recognition is the column where the invoice amouint is stored in the db.
mysql_cn.close() # Don't forget to close!



weeks = []

startw = str(raw_input("Enter the start week: "))
endw = str(raw_input("Enter the end week: "))
weeks.append(startw)
weeks.append(endw)

               
print "Fetching data from week " + startw + " to " + endw +" .... "

# Database types does not always get interpreted correctly when importing into a dataframe,
# so I always cast the data that I'm about to use in the correct format.

df[["SF_Code"]] = df[["SF_Code"]].astype(str)
df[["Recognition"]] = df[["Recognition"]].astype(int)
df[["Week_No"]] = df[["Week_No"]].astype(str)

print "Formatting done...."

df = df[df["Week_No"].between(weeks[0],weeks[1],inclusive=True)]


total = df.groupby(["SF_Code", "SF_Name"]).sum().sort_values(by="Recognition",ascending=False) # makes a summary of all sales in the given time period.



max_level = total.max(axis=0)
cut_off = int(raw_input("What is the first level? "))
steps = int(raw_input("What is the level increment? "))
stop = int(max_level["Recognition"])

total = total.loc[total["Recognition"] >= cut_off]

##
##levels = []
##for i in range(cut_off,stop,steps):
##    levels.append(int(i))

summary = []

for row in total["Recognition"]:
    if row > cut_off + steps:
        entry = ((row - cut_off)/steps)+1 # the +1 add the first level! 
        summary.append(entry)
    else:
        summary.append(1)

total["summary"] = summary




book = load_workbook("my spreadsheet location\myspreadsheet.xlsx")
writer = pd.ExcelWriter("my spreadsheet location\myspreadsheet.xlsx"", engine = 'openpyxl')
writer.book = book
writer.sheets = dict((ws.title,ws) for ws in book.worksheets)

total.to_excel(writer, sheet_name = "Sheet1", index=True)

writer.save()



file = "my spreadsheet location\myspreadsheet.xlsx""
os.startfile(file)
    















