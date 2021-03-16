import pandas as pd
book1 = pd.read_excel(r'Book1.xlsx', sheet_name='Sheet1')
book2 = pd.read_excel(r'Book1.xlsx', sheet_name='Sheet2')
f3 = book1.merge(book2, on ="Gmail", how = "left")
print(f3)