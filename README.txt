*********************************** Final Project 'Professional Software and Career Practises' ***********************************

This project performs data transformations and creates visualisations to analyse NFL statistics. It includes data merging, validation 
and visualisation of key metrics.

This repository includes the code I used in order to transform and manipulate the data I collected to then
create figures I can use in my report to visualise the data.

***************** Repository *****************

[GitHub Repository](https://github.com/abdullahrizmi14/DAT5902_final_project-1)

***************** Prerequisites *****************

- Python 3.8 or later
- pip (Python package manager)

***************** Setup Instructions *****************

1. Clone the repository:
   ```bash
   git clone https://github.com/abdullahrizmi14/DAT5902_final_project-1
   cd DAT5902_final_project-1

2. Create a virtual environment:
    On Linux/Mac:
    python -m venv .venv
    source .venv/bin/activate

    On Windows:
    python -m venv .venv
    .venv\Scripts\activate

3. Install required packages:
    pip install -r requirements.txt


***************** Files in the Repository *****************

I will explain the purpose for the files in my repository:

code1.py:

    - The primary code file containing all functions for importing, transforming, and analysing the data.
    - Generates datasets and figures used in the analysis.

unit_tests.pyy

    - A collection of unit tests for testing independent functions in the project.
    - Ensures critical components of the project function as expected.

requirements.txt:

    - A list of Python libraries and their versions required to run the project.
    - Generated using the command: 'pip freeze > requirements.txt'

data/:

    - Folder which contains datasets I used or created:

        - The_Data_Landscape_Project_Stats_Macro.xlsx: 

            The initial dataset containing multiple tables with individual statistics in one workbook.
        
        - transformation_workbook.xlsx:

            A playground dataset created during the transformation process. This file contains the merged and manipulated data used for
            testing methods before finalising.

        - final_data.csv:

            The final cleaned and transformed dataset saved as a CSV file. This file is used to generate the figures for the analysis.

Figures/:

    - Contains all the visualisations created using final_data.csv in code1.py.
    - These figures are used for analysis and reporting.


***************** Feautures in this project *****************

Data Transformation:

    - Merges multiple datasets into one cohesive structure.
    - Cleans and prepares data for analysis.

Visualization:

    - Generates insightful visualisations to analyse trends and statistics in NFL data.
    - Includes figures such as rushing vs. passing play percentages, win percentages, and touchdown comparisons.

Testing:

    - Includes a suite of unit tests to ensure the reliability and accuracy of key functions.


***************** Additional Notes *****************

Ensure the The_Data_Landscape_Project_Stats_Macro.xlsx file is present in the data/ directory for the project to run.

Figures will be saved in the Figures/ directory upon execution of the code.
