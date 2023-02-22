# Find all files within one mother folder containing a specified keyword 
# in their names and copy all to a new folder with no overlapping of the same file
# Liang Jingyi chez Sade Telecom created in 2021

import os
import shutil

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

dir_src = "C:/Users/liang.jingyi/Documents/CALCUL DE CHARGE/LCN/31286"
dir_new = "C:/Users/liang.jingyi/Documents/test_new"
keyword = "FicheAppui"

select_and_copy(dir_src, dir_new, keyword)


