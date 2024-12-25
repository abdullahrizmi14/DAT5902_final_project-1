import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl as xl      # To install this library 'pip install openpyxl'
import os
from adjustText import adjust_text


## Functions Used ##

def select_sheet(workbook, name):      # Display the first few rows of the sheet from excel workbook
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

def list_sheets(workbook):
    sheetNames = workbook.sheetnames
    for name in sheetNames:
        print(name)


## Loading in original workbook ##
data_xl = xl.load_workbook('The_Data_Landscape_Project_Stats_Macro.xlsx')

## Checking sheets for original workbook ##
print('---------------------------------- Original columns -----------------------------------------------')
list_sheets(data_xl)

## Deleting Merged data sheet and creating a new one ##
data_xl.create_sheet('merged_data_python')
del data_xl['Merged Data']

## Copying in team names to new merged sheet ## 
final_data_sheet = data_xl['merged_data_python']

for i in range(1,34):
    for j in range(1,2):
        read_value = data_xl.worksheets[0].cell(row = i, column = j) 
        final_data_sheet.cell(row = i, column = j).value = read_value.value

## Checking sheets of new workbook ##
print('----------------------------------- New columns ----------------------------------------------------')
list_sheets(data_xl)               ## checking to make sure that new sheet and 1st column are implemented

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
select_sheet(tf,'merged_data_python')    ## Checking head of sheet 


## Getting sheet names to make column headers ##
tf_shts = tf.sheetnames
headers = tf_shts[:-1]
last_sheet_name = tf_shts[-1]
last_sheet_tf = tf[last_sheet_name]

start_col = 2
start_row = 1

for col_index, header in enumerate(headers,start=start_col):
    print(f"Writing header '{header}' to column {col_index}")
    cell = last_sheet_tf.cell(row=start_row, column=col_index)
    cell.value = header

tf.save(file_name)




## Vlookup ##




























# for sheet_index,sheet in enumerate(headers,start=start_col):
#     for row in range(2, last_sheet_tf.max_row + 1):
#         last_sheet_tf.cell(row=row,column=sheet_index).value = (
#             f"=INDEX('{sheet}'!B:B, MATCH(A{row}, '{sheet}'!A:A, 0))"
#         )
# tf.save('transformation_workbook.xlsx')
# print("Workbook with fomulas saved")




# # Iterate through rows in the target sheet (starting from row 2 to skip the header)
# for row in range(2, last_sheet_tf.max_row + 1):
#     search_value = last_sheet_tf.cell(row=row, column=1).value  # Value to look up in column A
    
#     if search_value:  # Ensure the value is not empty
#         # Iterate through source sheets to find the match
#         for sheet_name in headers:
#             source_sheet = tf[sheet_name]
            
#             # Search for the matching value in column A of the source sheet
#             for source_row in range(2, source_sheet.max_row + 1):  # Assuming headers are in row 1
#                 if source_sheet.cell(row=source_row, column=1).value == search_value:
#                     # Retrieve the value from column B (or any specified column)
#                     matched_value = source_sheet.cell(row=source_row, column=2).value
                    
#                     # Write the retrieved value into the target sheet's column B
#                     last_sheet_tf.cell(row=row, column=2).value = matched_value
#                     break  # Stop searching once a match is found

# # Save the updated workbook
# tf.save('transformation_workbook_values.xlsx')
























# last_sheet = tf[last_sheet_name]

# Replace formulas with their computed values
# for row_index, row in enumerate(last_sheet_tf.iter_rows(values_only=True), start=1):
#     for col_index, value in enumerate(row, start=1):
#         last_sheet_tf.cell(row=row_index, column=col_index, value=value)

# # Save the workbook with updated values
# tf.save('transformation_workbook.xlsx')
# print("Workbook updated with values only.")





## Copy and Paste values ##
# tf = xl.load_workbook('transformation_workbook.xlsx', data_only=True)


# # Replace formulas with their computed values
# for row_index, row in enumerate(last_sheet_tf.iter_rows(values_only=True), start=1):
#     for col_index, value in enumerate(row, start=1):
#         last_sheet_tf.cell(row=row_index, column=col_index, value=value)

# # Save the workbook with updated values
# tf.save('transformation_workbook.xlsx')


# df=pd.read_excel('transformation_workbook.xlsx', sheet_name='merged_data_python')












# final_data_xl = xl.Workbook()
# new_sheet = final_data_xl.active
# new_sheet.title = last_sheet_name

# for row in last_sheet_tf.iter_rows(values_only=True):
#     new_sheet.append(row)

# final_data_xl.save('final_data.xlsx')
# print(f"Values from the last sheet '{last_sheet_name}' saved successfully.")

# select_sheet(tf,'merged_data_python') 


## Loading the Final dataset to CSV ##
# final_excel = 'transformation_workbook.xlsx'
# final__excel_sheet = 'merged_data_python'

# df = pd.read_excel(final_excel,sheet_name = final__excel_sheet)

# csv_name = 'final_data.csv'
# df.to_csv(csv_name, index=False)
# print(f"Sheet '{final__excel_sheet}' saved as '{csv_name}'.")

