import os
import sys
import openpyxl
import re




path2 = "C:\\Users\\raman\\Statistics\\"
path1 = "C:\\Users\\raman\\Moodle\\"

main_folder = []
edit_folder = []


os.chdir(path2)

files_in_statistics = os.listdir('.')

for files in range(0, len(files_in_statistics)):
    file_name = files_in_statistics[files]
    split_file_name = file_name.split(".")
    for split_name in split_file_name:
        if len(split_name) == 8 or len(split_name) == 10:
            edit_split_name = split_name[0:(len(split_name)-1)]
            main_folder.append(edit_split_name)
        elif len(split_name) == 7 or len(split_name) == 9:
            main_folder.append(split_name)

print(main_folder)


os.chdir(path1)


moodle_files = os.listdir('.')
print(len(moodle_files))

for moodle in range(0, len(moodle_files)):
    moodle_file_name = moodle_files[moodle]
    file_name = re.split("-| ", moodle_file_name)
    
    for splitted_file in file_name:
        if len(splitted_file) == 7 or len(splitted_file) == 9:
            if splitted_file not in main_folder:
                print(splitted_file)




print(edit_folder)