import pandas as pd
from openpyxl import load_workbook
# empty data frame is created
df_total = pd.DataFrame()
# all sheets are taken as a list
sheets = ['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4']
# from user email address is taken
yin = input("Enter Official Email Address")
# for loop for reading all the data in datasheets and searching for the data we need
for sheet in sheets:
    # loop through sheets inside an Excel file
    y = pd.read_excel('Excel.xlsx', sheet_name=sheet)
    df1 = pd.DataFrame(y[y['Official Email Address'] == yin], columns=y.columns)
    # here we are joining all the data of that particular person
    df_total = df_total.join(df1, how='outer', lsuffix='left', rsuffix='right')
# taking the path of the excel sheet
path = "Excel.xlsx"
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine='openpyxl')
writer.book = book
# searching for the master sheet in all sheets
if 'MasterSheet' in book.sheetnames:
    ref = book['MasterSheet']
    # removing the previous data in the sheet
    book.remove(ref)
    # writing the data into the master sheet
df_total.to_excel(writer, sheet_name='MasterSheet')
# saving  the data in the master sheet
writer.save()
writer.close()