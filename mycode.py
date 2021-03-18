import pandas as pd
from openpyxl import load_workbook
n=int(input("enter ps number"))
sheet_name=['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4']
s=pd.read_excel('Excel.xlsx', sheet_name)
y=s['Sheet1']
y=y[y['PS_NUMBER'] == n]
df = pd.DataFrame(y, columns=['PS_NUMBER','Display Name','Official Email Address'])

for i in s.keys():
    x=s[i]
    t=x[x['PS_NUMBER']==n]
    col = x.columns
    for j in col:
        df[j]=t[j]
#df.to_excel('pythonexcel1.xlsx',sheet_name='master',index=False)
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
df.to_excel(writer, sheet_name='MasterSheet')
# saving  the data in the master sheet
writer.save()
writer.close()