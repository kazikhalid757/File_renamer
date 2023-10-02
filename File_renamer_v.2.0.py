#!/usr/bin/env python
# coding: utf-8

# In[39]:


import os
import re 
import glob
import openpyxl


# In[79]:


def collecting_sub_folder_data():
    
    # Function to create an Excel file with subfolder names recursively
    def create_excel_file(parent_folder, worksheet, start_row=2):
        for root, dirs, files in os.walk(parent_folder):
            for dir_name in dirs:
                worksheet.cell(row=start_row, column=1, value=os.path.join(root, dir_name))
                start_row += 1

        # Save the Excel file
        excel_file = os.path.join(parent_folder, "subfolders.xlsx")
        worksheet.parent.save(excel_file)

        return excel_file

    # Input the parent folder path
    parent_folder = input("Enter the path of the parent folder: ")

    # Create the Excel file with subfolder names recursively
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Subfolders"

    # Add headers
    worksheet.cell(row=1, column=1, value="Subfolder Name")
    worksheet.cell(row=1, column=2, value="MAIN")
    worksheet.cell(row=1, column=3, value="asin_number")

    excel_file = create_excel_file(parent_folder, worksheet)

    print(f"Excel file 'subfolders.xlsx' created in '{parent_folder}'")
    print('\n')
    print("Please fill in the ASIN and MAIN columns as needed.")
    print('\n')
    print("Once you've filled in the Excel file, run the utility again to rename the files.")
    print('\n')


# In[80]:


def file_rename():
    # Function to read the Excel file and return folder paths and ASIN numbers
    def read_excel_file(excel_file):
        folder_data = []

        if not excel_file.endswith('.xlsx'):
            print("Invalid Excel file extension. Please provide a .xlsx file.")
            return folder_data

        if not os.path.isfile(excel_file):
            print("Excel file does not exist. Please provide a valid file path.")
            return folder_data

        try:
            workbook = openpyxl.load_workbook(excel_file)
            worksheet = workbook.active

            for row in worksheet.iter_rows(min_row=2, values_only=True):
                if len(row) >= 3:
                    folder_path, main, asin_number = row[:3]  # Extract folder path, main, and ASIN number
                    folder_data.append((folder_path, main, asin_number))
                else:
                    print("Skipping row with insufficient data:", row)

        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")

        return folder_data

    # Function to rename files within a folder based on the provided condition
    def rename_files(folder_path, asin_number):
        if not os.path.exists(folder_path):
            print(f"Folder not found: {folder_path}")
            return

        pt_number = 1  # Initialize the part number

        for root, _, files in os.walk(folder_path):
            for filename in files:
                base, ext = os.path.splitext(filename)

                if base.endswith("MAIN_1"):
                    new_filename = f"{asin_number}.MAIN{ext}"
                elif base.endswith("_1"):
                    new_filename = f"{asin_number}.MAIN{ext}"
                else:
                    new_filename = f"{asin_number}.PT{pt_number:02d}{ext}"
                    pt_number += 1  # Increment part number for non-main files

                src_file_path = os.path.join(root, filename)
                dst_file_path = os.path.join(root, new_filename)

                if not os.path.exists(dst_file_path):
                    os.rename(src_file_path, dst_file_path)
                    print(f"Renamed: {filename} -> {new_filename}")

    # Input the path of the Excel file
    excel_file = input("Enter the full path of the Excel file (including filename and extension): ")

    folder_data = read_excel_file(excel_file)

    if folder_data:
        for folder_path, main, asin_number in folder_data:
            print(f"Processing folder: {folder_path}, ASIN: {asin_number}")
            rename_files(folder_path, asin_number)


# In[83]:


print("Enter 1 for collecting sub-folder data. ")
print("Enter 2 for file rename. ")
num=int(input("Enter Number"))
if num==1:
    collecting_sub_folder_data()
    
else:
    file_rename()
    


# In[ ]:




