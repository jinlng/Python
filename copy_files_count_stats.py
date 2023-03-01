import os
import shutil
import pandas as pd

folder_org = "C:/Users/liang.jingyi/Documents/CALCUL DE CHARGE/LCN/31297"
folder_new = "C:/Users/liang.jingyi/Documents/test_new"
keyword = "FicheAppui"

# Read the excel file
recap_file = os.path.join(folder_org, "recap.xlsx")
df = pd.read_excel(recap_file)

# Extract the values in column J and convert them into a list
numbers = df.iloc[:, 9].astype(str).tolist()

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
                    print(f"Copied file: {filename}")
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