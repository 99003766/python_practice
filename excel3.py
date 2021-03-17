import pandas as pd
book1 = pd.read_excel(r'Excel.xlsx', sheet_name='Sheet1')
book2 = pd.read_excel(r'Excel.xlsx', sheet_name='Sheet2')
f3 = book1.combine(book2, on="PS_NUMBER", how="left")
print(f3)