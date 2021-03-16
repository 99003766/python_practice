import pandas as pd
# import numpy as np
from openpyxl import load_workbook
# import openpyxl
excel_file = pd.ExcelFile('crime.xlsx')

df_mini_first = pd.read_excel(excel_file, sheet_name=0)
df_mini_second = pd.read_excel(excel_file, sheet_name=1)
df_mini_third = pd.read_excel(excel_file, sheet_name=2)
df_mini_fourth = pd.read_excel(excel_file, sheet_name=3)
df_mini_fifth = pd.read_excel(excel_file, sheet_name=4)

f3 = df_mini_first.merge(df_mini_second,  on="Area_Name", how="left")
f4 = f3.merge(df_mini_third, on="Area_Name", how="left")
f5 = f4.merge(df_mini_fourth, on="Area_Name", how="left")
f6 = f5.merge(df_mini_fifth, on="Area_Name", how="left")
Input = input("Area name")
df_Input = pd.DataFrame(f6, columns=['Area_Name', 'Year(2001)', 'Group_Name(2001)', 'Sub_Group_Name(2001)', 'Cases_Property_Recovered(2001)', 'Cases_Property_Stolen(2001)', 'Value_of_Property_Recovered(2001)', 'Value_of_Property_Stolen(2001)', 'Population(2001)',    'average_age(2001)',
                                     'Year(2002)', 'Group_Name(2002)', 'Sub_Group_Name(2002)', 'Cases_Property_Recovered(2002)', 'Cases_Property_Stolen(2002)', 'Value_of_Property_Recovered(2002)', 'Value_of_Property_Stolen(2002)', 'Population(2002)', 'average_age(2002)',
                                     'Year(2003)', 'Group_Name(2003)', 'Sub_Group_Name(2003)', 'Cases_Property_Recovered(2003)', 'Cases_Property_Stolen(2003)', 'Value_of_Property_Recovered(2003)', 'Value_of_Property_Stolen(2003)', 'Population(2003)', 'average_age(2005)',
                                     'Year(2004)', 'Group_Name(2004)', 'Sub_Group_Name(2004)', 'Cases_Property_Recovered(2004)', 'Cases_Property_Stolen(2004)', 'Value_of_Property_Recovered(2004)', 'Value_of_Property_Stolen(2004)', 'Population(2004)', 'average_age(2004)',
                                     'Year(2005)', 'Group_Name(2005)', 'Sub_Group_Name(2005)', 'Cases_Property_Recovered(2005)', 'Cases_Property_Stolen(2005)', 'Value_of_Property_Recovered(2005)', 'Value_of_Property_Stolen(2005)', 'Population(2005)', 'average_age(2005)'])
# print(pf_variable)
df_Input.set_index('Area_Name', inplace=True)
result = df_Input.loc[Input]
print(result)
path = "crime.xlsx"
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine='openpyxl')
writer.book = book
if 'Mastersheet' in book.sheetnames:
    ref = book['Mastersheet']
    book.remove(ref)
result.to_excel(writer, sheet_name='Mastersheet')
writer.save()
writer.close()
