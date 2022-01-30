# - *- coding: utf- 8 - *-
import pandas as pd
import os
import glob


try:

        mainPath = os.getcwd()
        os.chdir(mainPath + '/Files for appending')
        df =  pd.DataFrame

        for file in list(glob.glob('*.xlsx*')):
                
                df2 = pd.read_excel(file)

                if not df.empty:    
                        df = df.append( df2 )
                else:
                        df = df2

        excel_writer = pd.ExcelWriter(mainPath + '/Summary_appended.xlsx')
        df.to_excel(excel_writer)
        excel_writer.save()

except:

        with open (mainPath + '/error.txt','w' ) as f:
                f.write('The email column should be renamed to Email and the grade column to Grade')
                print('error')

    


