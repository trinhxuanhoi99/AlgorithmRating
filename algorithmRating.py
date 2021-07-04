import pandas as pd
import xlrd
import numpy as np
import openpyxl
df=pd.read_excel('testdataa.xlsx',sheet_name='itemdata')


c4=df['star4'].sum()
c5=df['star5'].sum()
v1=df.star1+df.star2+df.star3+df.star4+df.star5
R1=df.avgrating
m1=min(v1)
c1=c4+c5
df['Analyst2']= (v1/(v1+m1))*R1+(m1/(v1+m1))*c1
df[['Analyst2']]=df[['Analyst2']].fillna(value=0)
print(df['Analyst2'])
# df.to_excel(r'C:\Users\DELL\Desktop\export_dataframe2.xlsx',index=False,header=True)
