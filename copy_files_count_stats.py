"""
Title: Copy Files with Keyword
Description:
    This code is designed to copy files from a source folder to a destination folder based on a specified keyword and certain conditions.
    It operates by utilizing a recap.xlsx file, which you need to create and fill in. Please note that in the recap.xlsx file,
    the first line should be left empty to ensure that the first value is copied correctly to the new folder.
Author: Jingyi LIANG
Date: May 8, 2023
License: This code is the property of Jingyi LIANG. Unauthorized use or distribution is strictly prohibited.
"""

import os
import shutil
import pandas as pd

# Define the path to the input Excel file
folder_org = "C:/Users/liang.jingyi/Documents/CALCUL DE CHARGE/PAS/31321/D2"
folder_new = "C:/Users/liang.jingyi/Documents/test_new"
keyword = "FicheAppui"

# Read the excel file
recap_file = os.path.join(folder_org, "recap.xlsx")
df = pd.read_excel(recap_file)

# Extract the values in column J + K and convert them into a list
numbers = [str(int(x)) for x in df.iloc[:, 9:11].values.flatten() if pd.notna(x) and str(x).strip()]

# Count the number of times the keyword appears in column H
numbers_KO = df.iloc[:, 7].str.contains("REMPLACEMENT", na=False).sum()


# Define functions
def count_files(folder_org, folder_new, keyword):
    count = 0
    for dirpath, dirnames, filenames in os.walk(folder_org):
        for filename in filenames:
            if keyword in filename:
                try:
                    count += 1
                except (shutil.SameFileError, PermissionError):
                    pass
                except Exception as e:
                    print(f"Error occurred while copying file: {e}")
        for dirname in dirnames:
            count_files(os.path.join(dirpath, dirname), folder_new, keyword)
    return count


def select_and_copy(folder_org, folder_new, keyword, numbers):
    copied_files = set()
    for dirpath, dirnames, filenames in os.walk(folder_org):
        for filename in filenames:
            if keyword in filename and any(str(num) in filename for num in numbers):
                try:
                    shutil.copy(os.path.join(dirpath, filename), os.path.join(folder_new, filename))
                    copied_files.add(filename)
                    #print(f"Copied file: {filename}")
                except (shutil.SameFileError, PermissionError):
                    pass
                except Exception as e:
                    print(f"Error occurred while copying file: {e}")
        for dirname in dirnames:
            copied_files.update(select_and_copy(os.path.join(dirpath, dirname), folder_new, keyword, numbers))
    return copied_files

count_files = count_files(folder_org, folder_new, keyword)
copied_files = select_and_copy(folder_org, folder_new, keyword, numbers)
print(f"Number of files matching '{keyword}': {count_files}")
print(f"Number of files copied to a new folder: {len(copied_files)}")
print(f"Number of KO: {numbers_KO}")