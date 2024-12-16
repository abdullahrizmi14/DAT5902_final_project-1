import pandas as pd
import numpy as np
import openpyxl as xl      # To install this library 'pip install openpyxl'


## Loading workbooks and checking sheets ##
data = xl.load_workbook('The_Data_Landscape_Project_Stats_Macro.xlsx')

sheetNames = data.sheetnames


print('---------------------------------- Original columns -----------------------------------------------')
for name in sheetNames:
    print(name)



## Deleting Merged data sheet and creating a new one ##

del data['Merged Data']

data.create_sheet('merged_data_python')

sheetNames = data.sheetnames


print('----------------------------------- New columns ----------------------------------------------------')
for name in sheetNames: ## checking to make sure that new names are implemented
    print(name)
