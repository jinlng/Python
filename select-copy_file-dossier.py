# Find all files within one mother folder containing a specified keyword 
# in their names and copy all to a new folder with no overlapping of the same file
# Liang Jingyi chez Sade Telecom created in 2021

import os
import shutil
import pandas as pd

def select_and_copy(folder_org, folder_new, keyword):
    for dirpath, dirnames, filenames in os.walk(folder_org):
        for filename in filenames:
            if keyword in filename:
                try:
                    shutil.copy(os.path.join(dirpath, filename), os.path.join(folder_new, filename))
                except (shutil.SameFileError, PermissionError):
                    pass
                except Exception as e:
                    print(f"Error occurred while copying file: {e}")
        for dirname in dirnames:
            select_and_copy(os.path.join(dirpath, dirname), folder_new, keyword)

dir_src = "C:/Users/liang.jingyi/Documents/CALCUL DE CHARGE/LCN/31295"
dir_new = "C:/Users/liang.jingyi/Documents/test_new"
keyword = "FicheAppui"

# Read the excel file
recap_file = os.path.join(dir_src, "recap.xlsx")
df = pd.read_excel(recap_file)

# Extract the values in column J and convert them into a list
numbers = df['J'].tolist()

# Call the select_and_copy function with the list of numbers
for number in numbers:
    select_and_copy(dir_src, dir_new, str(number))

count = select_and_copy(dir_src, dir_new, keyword)
print(f"Number of files containing '{keyword}': {count}")
