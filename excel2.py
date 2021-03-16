import pandas as pd
book1 = pd.read_excel(r'Book1.xlsx', sheet_name='Sheet1')
book2 = pd.read_excel(r'Book1.xlsx', sheet_name='Sheet2')
pd.wr = pd.concat([book1, book2])
#print(pd.ExcelWriter('Book1.xlsx', sheet_name='Sheet3') == pd.wr)
print(pd.wr)