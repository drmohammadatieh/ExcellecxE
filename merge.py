# - *- coding: utf- 8 - *-
import pandas as pd
import os
import glob

# try:

mainPath = os.getcwd()
os.chdir(mainPath + '/Files for merging')
# os.chdir(mainPath + '\Tests') # In Windows

df =  pd.DataFrame()

print(df)
for file in list(glob.glob('*.xlsx*')):
    df2 = pd.read_excel(file)[['ID','Grade']]

    if not df.empty:    
            df = pd. merge( df,df2 , on='ID',how='outer' )
    else:
        df = df2

index = 0
counter = 1
for column in df.columns:
    if "Grade" in column:
        df.columns.values[index] = 'Test ' + str(counter)
        counter +=1
    index += 1

df = df.set_index(drop=True)
excel_writer = pd.ExcelWriter(mainPath + '/Summary_merged.xlsx')
df.to_excel(excel_writer)

# worksheet=excel_writer.sheets['Sheet1']
# header_list = df.columns.values.tolist()

excel_writer.save()

# except:

#     with open (mainPath + '/error.txt','w' ) as f:
#         f.write('The email column should be renamed to Email and the grade column to Grade')
#         print('error')

    


