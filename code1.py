import pandas as pd
import numpy as np
import openpyxl as xl      # To install this library 'pip install openpyxl'
import os


## Functions Used ##

def select_sheet_xl(workbook, name):      # Display the first few rows of the sheet from excel workbook
    """
    Display the first few rows of the specified sheet from an openpyxl workbook.
    
    Args:
        workbook: An openpyxl Workbook object.
        name: The name of the sheet to display.
    """
    if name in workbook.sheetnames:
        df = workbook[name]
        print("Data from " + name + ":")
        for row in df.iter_rows(min_row=1, max_row=5, values_only=True):  # Display the first few rows of the DataFrame
            print(row)    
    else:
        print(name + "sheet not found.")


# def specific_value_pd(workbook, sheetnm, row, column):
#     data = pd.read_excel(workbook, sheet_name = sheetnm)

#     row_index = row - 1
#     column_index = column - 1

#     cell_value = df.iloc[row_index, column_index]
#     print(cell_value)


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

## Checking sheets of new workbook##
print('----------------------------------- New columns ----------------------------------------------------')
list_sheets_xl(data_xl)               ## checking to make sure that new sheet and 1st column are implemented

## Save the workbook ##
file_name = 'transformation_workbook.xlsx'

if os.path.exists(file_name):
    os.remove(file_name)
    print(f"Existing file'{file_name}' deleted.")
else:
    print(f"No existing file found with name'{file_name}'.")

data_xl.save(file_name)
print(f"New file '{file_name}' created successfully.")



## Load new workbook for transformations and check head ##
tf = xl.load_workbook(file_name)
#select_sheet_xl(tf,'merged_data_python')    ## Checking head of sheet 


## Getting sheet names to make column headers ##
tf_shts = tf.sheetnames
headers = tf_shts[:-1]
last_sheet_name = tf_shts[-1]
last_sheet_tf = tf[last_sheet_name]

start_col = 2
start_row = 1


for col_index, header in enumerate(headers,start=start_col):
    cell = last_sheet_tf.cell(row=start_row, column=col_index)
    cell.value = header


tf.save('transformation_workbook.xlsx')




## Vlookup ##
# last_sheet_tf['B2'] = "=INDEX('Points Per Game'!B:B, MATCH(A2, 'Points Per Game'!A:A, 0))"

# for row in range(2, last_sheet_tf.max_row + 1):
#     last_sheet_tf[f"B{row}"] = f"=INDEX('Points Per Game'!B:B, MATCH(A{row}, 'Points Per Game'!A:A, 0))"

# tf.save('transformation_workbook.xlsx')


for sheet_index,sheet in enumerate(tf_shts,start=start_col):
    
    for row in range(2, last_sheet_tf.max_row + 1):
        last_sheet_tf.cell(row=row,column=sheet_index).value = (
            f"=INDEX('{sheet}'!B:B, MATCH(A{row}, '{sheet}'!A:A, 0))"
        )

tf.save('transformation_workbook.xlsx')












# # Access the first and last sheets
# first_sheet = tf.worksheets[0]  # Assuming the first sheet is the source
# last_sheet = tf.worksheets[-1]  # Assuming the last sheet is the target

# # Iterate over rows in the last sheet (starting from the second row to skip the header)
# for row in range(2, last_sheet.max_row + 1):  # row=2 to skip the header
#     search_value = last_sheet.cell(row=row, column=1).value  # Get the value in the first column
    
#     if search_value:  # Ensure the value is not empty
#         # Look for the matching value in the first column of the first sheet
#         for source_row in range(2, first_sheet.max_row + 1):  # Assuming the source also has a header
#             source_value = first_sheet.cell(row=source_row, column=1).value
            
#             if search_value == source_value:  # Match found
#                 # Copy data from the source sheet to the target sheet
#                 for col in range(2, first_sheet.max_column + 1):  # Start from column 2 in source
#                     value_to_copy = first_sheet.cell(row=source_row, column=col).value
#                     last_sheet.cell(row=row, column=col).value = value_to_copy  # Write into last sheet
#                 break  # Stop searching once a match is found

# Save the updated workbook
#tf.save('transformation_workbook.xlsx')
# print("VLookup operation completed successfully!")
