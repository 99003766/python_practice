import pandas as pd
from openpyxl import load_workbook
# empty data frame is created
df_total = pd.DataFrame()
res=[]
# all sheets are taken as a list
sheets = ['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4']
# from user email address is taken
yin = input("Enter Official Email Address")
# for loop for reading all the data in datasheets and searching for the data we need
for i in sheets:
    # loop through sheets inside an Excel file
    y = pd.read_excel('Excel.xlsx', sheet_name=i)
    number = i
    for i in range(i.nrows):
        if (i.row_values(i)[0] == PS_NUMBER and sheet.row_values(i)[1] == PsNo and sheet.row_values(i)[2] == EmailId):
            print(i.row_values(i))
    if number == 0:
        for i in i.row_values(i):
            if i not in res:
            res.append(i)
    else:
        print(i.row_values(i)[3:])
    for j in i.row_values(i)[3:]:
        res.append(j)
    # print(res)