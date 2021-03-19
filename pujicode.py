import pandas as pd
h=int(input("enter ps number"))
z=pd.read_excel('Excel.xlsx',sheet_name=['Sheet1','Sheet2','Sheet3','Sheet4'])
y=z['Sheet1']
y=y[y['PS_NUMBER']==h]
df = pd.DataFrame(y, columns = ['PS_NUMBER','Display Name','Official Email Address'])

for i in z.keys():
    x=z[i]
    t=x[x['PS_NUMBER']==h]
    col = x.columns
    for j in col:
        df[j]=t[j]
df.to_excel('pythonexcel1.xlsx',sheet_name='master', index=False)