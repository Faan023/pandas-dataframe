import MySQLdb
import pandas as pd
from tabulate import tabulate
from openpyxl import load_workbook
import os


mysql_cn = MySQLdb.connect(host="my_host_name",
                           port=3306,user="my_super_user",passwd="my_password",
                           db="my_db")
print " MySQL connection made...."

df = pd.read_sql( "select Invoice_Number, Stock_Code, Stock_Description, Quantity from stock_movement2016;",con=mysql_cn)
print " Dataframe created..."
print " " + str(len(df)) + " lines of data imported..."

mysql_cn.close()


with open("silverton.txt") as read_silverton:
    content_silverton = read_silverton.read().splitlines()

with open("pplace.txt") as read_pplace:
    content_pp = read_pplace.read().splitlines()

with open("hammans.txt") as read_hammans:
    content_hammans = read_hammans.read().splitlines()


            
print " Invoices read in...."



df[["Invoice_Number"]] = df[["Invoice_Number"]].astype(str)
df[["Quantity"]] = df[["Quantity"]].astype(int)

print " Formatting done...."

stock_silverton = df[df["Invoice_Number"].isin(content_silverton)]
stock_pp = df[df["Invoice_Number"].isin(content_pp)]
stock_hammans = df[df["Invoice_Number"].isin(content_hammans)]
print " Filter applied..."

final_silverton = stock_silverton.groupby(["Stock_Code","Stock_Description"]).sum()
final_pp = stock_pp.groupby(["Stock_Code","Stock_Description"]).sum()
final_hammans = stock_hammans.groupby(["Stock_Code","Stock_Description"]).sum()

total= [final_silverton,final_pp,final_hammans]

# for i in total:
#     print "The summary for " + i + " is: "
#     print tabulate(i,tablefmt="grid")

# print "The summary for " + whouse + " is: "
# print tabulate(final,tablefmt="grid")

# ex = ()
# while ex != "y":
#     ex = raw_input("Press 'y' to export, any other key to quit: ")
book = load_workbook("cc_export.xlsx")
writer = pd.ExcelWriter("cc_export.xlsx", engine = 'openpyxl')
writer.book = book
writer.sheets = dict((ws.title,ws) for ws in book.worksheets)

final_silverton.to_excel(writer, sheet_name = "silverton", index=True)  
final_pp.to_excel(writer, sheet_name = "pplace", index=True) 
final_hammans.to_excel(writer, sheet_name = "hammans", index=True)  

writer.save()



file = "cc_export.xlsx"
os.startfile(file)




    

    






           


        
        


      

    
        
    





