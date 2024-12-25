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

def check_if_exists_then_create(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f"Existing file'{file_path}' deleted.")
    else:
        print(f"No existing file found with name'{file_path}'.")


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
check_if_exists_then_create(file_name)
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

# Save
tf.save(file_name)


## Vlookup ##

# Loading Workbook and sheet names
main_wb = pd.read_excel(file_name, sheet_name=None)
sheetNames_wb = list(main_wb.keys())

# Loading last sheet and columns
main_data = pd.read_excel(file_name, sheet_name='merged_data_python')  # Main data
columns_md = list(main_data.columns)


for i, sheet_name in enumerate(sheetNames_wb):
    if i >= len(columns_md)-1:
        print("Not enough columns in main_data to update. Skipping remaining sheets.")
        break

    # Iterating through sheets
    lookup_data = pd.read_excel(file_name, sheet_name=sheet_name)  # Lookup data
    columns_ld = list(lookup_data.columns)

    # Merge on common key 'Team'
    merged_data = main_data.merge(lookup_data[columns_ld[:2]], on='Team', how='left')

    main_data[columns_md[i + 1]] = merged_data[columns_ld[1]]

# Save to CSV
main_data.to_csv('final_data.csv', index=False)
print("VLOOKUP operation completed successfully.")



#######################################################################################################################

## Analysis and Figures ##

## Loading Data ##
df = pd.read_csv('final_data.csv')
df

## Figure 1 'Rushing and Passing Play %'s ##

## Storing data for Rushing Play % and Passing Play %
rushing_03 = df['Rushing Play % (2003)']
rushing_13 = df['Rushing Play % (2013)']
rushing_23 = df['Rushing Play % (2023)']

passing_03 = df['Passing Play % (2003)']
passing_13 = df['Passing Play % (2013)']
passing_23 = df['Passing Play % (2023)']

# Create a figure
plt.figure(figsize=(10, 6))

# Plot KDE for Rushing Play %
sns.kdeplot(rushing_03, shade=True, label='Rushing (2003)', color='blue', bw_adjust=1.2)
sns.kdeplot(rushing_13, shade=True, label='Rushing (2013)', color='green', bw_adjust=1.2)
sns.kdeplot(rushing_23, shade=True, label='Rushing (2023)', color='red', bw_adjust=1.2)

# Plot KDE for Passing Play %
sns.kdeplot(passing_03, shade=True, label='Passing (2003)', color='purple', bw_adjust=1.2)
sns.kdeplot(passing_13, shade=True, label='Passing (2013)', color='orange', bw_adjust=1.2)
sns.kdeplot(passing_23, shade=True, label='Passing (2023)', color='brown', bw_adjust=1.2)

# Titles and labels
plt.xlabel('Percentage', fontsize = 15, labelpad=15, fontweight = 'bold')
plt.ylabel('Density', fontsize = 15, labelpad=15, fontweight = 'bold')
plt.title('Rushing and Passing Play % (2003, 2013, 2023)')

# Add legend
plt.legend(loc='upper right')


## Save the figure as an image file (e.g., PNG or JPG)
plt.savefig('play_percentage_distribution_20years.png', format='png', dpi=300)






