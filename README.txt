*********************************** Final Project 'Professional Software and Career Practises' ***********************************

This project performs data transformations and creates visualizations to analyze NFL statistics. It includes data merging, validation 
and visualization of key metrics.

This repository includes the code I used in order to transform and manipulate the data I collected to then
create figures I can use in my report to visualise the data.

***************** Prerequisites *****************

- Python 3.8 or later
- pip (Python package manager)

***************** Setup Instructions *****************

1. Clone the repository:
   ```bash
   git clone https://github.com/DAT5902_final_project-1.git
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

    - This is the main code file. Includes all my code on functions, importing, transforming the data.
      This is also the file the in which I create my datasets and figures in.

unit_tests.py:

    - The file which contains all my unit tests created. They test independent functions whcih can be tested
      independently

requirements.txt:

    - This is a file that circle.ci reads in order to test my unit tests, including the imported libraries I
      used. This was done by running 'pip freeze > requirements.txt' in the terminal of my environment

data/:

    - Folder which contains datasets I used or created:

        - The_Data_Landscape_Project_Stats_Macro.xlsx: 

            Data I began with, consists of many tables each with an individual statistic in one workbook.
        
        - transformation_workbook.xlsx:

            Data for my tranformations, this workbook was my 'playground'. I imported data from 'The_Data_Landscape_Project_Stats_Macro.xlsx' 
            by merging all the smaller datasets into one large one. This dataset was me trialing my methods before I saved it
            all to the final form.

        - final_data.csv:

            This csv file was my polished data after the data was cleaned and transformed. Saved in a CSV file.

Figures/:

    - This folder stored all the figures I created from 'final_data.csv' in 'code1.py' to use in my analysis





