# Jacob Hui - Clover Automation

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from methods import get_departments, item_department_dict

# import file path from clover
clover_folder = input(r"Input the folder path that you are using: ") 
clover_file = input(r"Input file name: ") + ".xlsx"
# print(clover_folder, clover_file)

# test file
# clover_folder = r"C:\Users\Ovvi\OneDrive\Desktop\Automation\Clover" 
# clover_file = r"\testclover.xlsx"
clover_file = "testclover.xlsx"

# sanity check
clover_path = clover_folder + clover_file
print(clover_folder)
print(clover_file)
print(clover_path)

# load worksheet
clover_wb = load_workbook(clover_path)
clover_ws = clover_wb.active
print(clover_wb.sheetnames)

# get departments
departments = get_departments(clover_wb)

# create dictionary of items with assigned department
items_dict = item_department_dict(clover_wb, departments)



# https://www.programiz.com/python-programming/nested-dictionary

# Item structure
    # item_name = {
    # department: "",
    # price: float(),
    # cost: float(),
    # barcode: ""
    # }

# categories sheet 
# get item
# get department 


# find item in items sheet 
# populate item attributes 

# store in new sheet
# save 
