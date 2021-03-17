import pandas as pd
df_total = pd.DataFrame()
sheets=['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4']
xin = input()
for sheet in sheets: # loop through sheets inside an Excel file
    x = pd.read_excel('Excel.xlsx', sheet_name=sheet)
    df1 = pd.DataFrame(x[x['Display Name']==xin], columns=x.columns)
    df_total = df_total.join(df1, how='outer', lsuffix='_left', rsuffix='_right')
df_total.to_excel('Excel2.xlsx', index=False)