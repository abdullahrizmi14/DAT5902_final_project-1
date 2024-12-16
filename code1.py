import pandas as pd
import numpy as np
import openpyxl as xl      # To install this library 'pip install openpyxl'

## Functions Used ##

def select_sheet_pd(name):
    if name in all_data_pd:
        df = all_data_pd[name]
        print("Data from " + name + ":")
        print(df.head())  # Display the first few rows of the DataFrame
    
    else:
        print(name + "sheet not found.")


## Loading in all workbooks needed ##
data_xl = xl.load_workbook('The_Data_Landscape_Project_Stats_Macro.xlsx')
all_data_pd = pd.read_excel('The_Data_Landscape_Project_Stats_Macro.xlsx', sheet_name = None)


## Loading workbooks and checking sheets ##
sheetNames = data_xl.sheetnames

print('---------------------------------- Original columns -----------------------------------------------')
for name in sheetNames:
    print(name)

## Deleting Merged data sheet and creating a new one ## 
del data_xl['Merged Data']
data_xl.create_sheet('merged_data_python')
sheetNames = data_xl.sheetnames

print('----------------------------------- New columns ----------------------------------------------------')
for name in sheetNames:                           ## checking to make sure that new names are implemented
    print(name)


## Copying in team names to new merged sheet ## 
final_data_sheet = data_xl['merged_data_python']



for i in range(1,34):
    for j in range(1,2):

        read_value = data_xl.worksheets[0].cell(row = i, column = j) 

        final_data_sheet.cell(row = i, column = j).value = read_value.value



print(final_data_sheet.cell(row = 33, column = 1).value)

print(data_xl['merged_data_python']['Team'])