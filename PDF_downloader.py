#### EDIT THE VALUES BELOW TO FIT THE USE CASE


# Specify the path and filename of the excel file containing the download links
# Can be done by selecting the excel file and then using Ctrl + Shift + C
# Remember to let the r before the string stay so that Python knows it's a raw string!
excel_file_path = r"C:\Users\Spac-23\Documents\Opgaver - Softwarespor\Uge 4\GRI_2017_2020 (1) (test version).xlsx"

# Specify which columns in the excel file to locate the download links
# Value should be a list, and each element can be either:
# 1: A combination of uppercase letters, such as B, AM, or BL, corrosponding to the columns in the excel sheet
# 2: Positive integers, such as 2, 14, or 145
# The columns are prioritized, the program will start by checking the first column and then go from there
download_link_columns = ["AL", "AM"] #TODO: ændres

# Specify a column in the excel file for the program to write whether ad PDF file has been downloaded or not
# Value can be either:
# 1: A combination of uppercase letters, such as B, AM, or BL, corrosponding to the columns in the excel sheet
# 2: A positive integer, such as 2, 14, or 145
# Note that a column should be created in the excel file for the purpose of this, if one is not already created
is_downloaded_column = "Is downloaded?" #TODO: ændres

# Specify whether the data in the excel sheet contains a header
# Value should be either True or False
contains_header = True #TODO: Fjernes?

# Specify the path and folder to place the downloaded files
# Can be done by selecting the folder and then using Ctrl + Shift + C
# Remember to let the r before the string stay so that Python knows it's a raw string!
destination_folder = r"C:\Users\Spac-23\Documents\Opgaver - Softwarespor\Uge 4\test download folder"

# Specify whether to limit the amount of downloads each time the program is run
# Value should be either True or False
restricted_mode = True

# Specify how many files to download when the program runs in restricted mode
# Only used when restricted_mode = True
# Value should be a positive integer, such as 5, 10, or 20
# For the sake of robustness, value should still be a positive integer even when restricted_mode = False
max_downloads = 10


#### DO NOT CHANGE ANYTHING BELOW THIS LINE UNLESS YOU WISH TO ALTER THE FUNCTIONALITY OF THE PROGRAM



## Imports
import pandas
import PyPDF2
from pathlib import Path
import shutil, os
import os.path
import urllib.request # Do we only need this one subfuntion or the whole library?
import glob
import requests
#TODO: Comment out some of these to test if it's used or not

## function definitions
# Function to translate column strings into indecies (should leave positive integers as is)
#translate_to_index([list or string or positive integer]): ...

# Function to download a PDF file given a download link and a destination folder

## Other variables

# Keep a list of indecies of which files have been downloaded
is_downloaded_indecies = []

#ownload_columns_indecies = translate_to_index(download_columns) # Maybe this should be in the part where the code runs?

## Running the code

# The code should not run if imported somewhere else
if __name__ == "__main__":
    # Load the excel file into a data frame
    dataframe = pandas.read_excel(excel_file_path)

    print(dataframe['BRnum'])
    print(dataframe['BRnum'][3])
    print(dataframe['Report Html Address'][18599-2])
    print(dataframe.shape[0])

    filename = str(destination_folder + "/test.pdf") #TODO: write comments and make sure it writes a proper filename
    response = requests.get(dataframe['Report Html Address'][18599-2])
    with open(filename,'wb') as file:
        file.write(response.content)

    # Define iteration variable and find the number of rows in the dataset
    downloads = 0 # Keep track of the number of successful downloads
    max = dataframe.shape[0]

    # Looping through the data frame
    # The loop should exit when there's no more data
    while i < max: #TODO: make for loop instead, also make number of successfull downloads a different variable


######## TIL TESTNING: Følgende adresse giver faktisk en PDF fil: https://www.tateandlyle.com/sites/default/files/2019-06/tatelyle-annual-report-2019-interactive-pdf.pdf
# Ligger som udgangspunkt i række 18599