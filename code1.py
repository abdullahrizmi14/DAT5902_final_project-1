import pandas as pd
import numpy as np
import openpyxl as xl      # To install this library 'pip install openpyxl'

## Functions Used ##

def select_sheet_pd(name):      # Display the first few rows of the sheet from excel workbook
    if name in all_data_pd:
        df = all_data_pd[name]
        print("Data from " + name + ":")
        print(df.head())  # Display the first few rows of the DataFrame
    else:
        print(name + "sheet not found.")


def specific_value_pd(workbook, sheetnm, row, column):
    data = pd.read_excel(workbook, sheet_name = sheetnm)

    row_index = row - 1
    column_index = column - 1

    cell_value = df.iloc[row_index, column_index]
    print(cell_value)


def list_sheets_xl(workbook):
    sheetNames = workbook.sheetnames
    for name in sheetNames:
        print(name)


## Loading in original workbook ##
data_xl = xl.load_workbook('The_Data_Landscape_Project_Stats_Macro.xlsx')

## Checking sheets for original workbook ##
print('---------------------------------- Original columns -----------------------------------------------')
list_sheets_xl(data_xl)

## Deleting Merged data sheet and creating a new one ##
data_xl.create_sheet('merged_data_python')
del data_xl['Merged Data']

## Copying in team names to new merged sheet ## 
final_data_sheet = data_xl['merged_data_python']

for i in range(1,34):
    for j in range(1,2):
        read_value = data_xl.worksheets[0].cell(row = i, column = j) 
        final_data_sheet.cell(row = i, column = j).value = read_value.value


print('----------------------------------- New columns ----------------------------------------------------')
list_sheets_xl(data_xl)               ## checking to make sure that new sheet is implemented


## Save the workbook and reload with pandas ##
data_xl.save('workbook_for_pandas.xlsx')

## Load new workbook with pandas ##
all_data_pd = pd.read_excel('workbook_for_pandas.xlsx', sheet_name = None)
select_sheet_pd('merged_data_python')    ## Checking head of sheet 










## Adding new column headers to new merged sheet
sheetNames_pandas = list(all_data_pd.keys())
headers = sheetNames_pandas[:-1]
print(headers)