# Jacob Hui - Clover Automation

from methods import get_departments, item_department_dict, initialItemIstance, Ovvi
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

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

items = initialItemIstance(clover_wb, departments)
print('-')
# print(items[0].itemName)
# print(items[0].itemDepartment)

index = 0
for item in items:
    if item.itemName == 'Kiwi':
        print(item.itemName)
        print(item.itemDepartment)
        print(index)
    index += 1

# categories sheet 
# get item
# get department 

# find item in items sheet 
ws = clover_wb["Items"] # select categories ws
for row in ws.iter_rows():
    print(row)
    

# populate item attributes 

# store in new sheet
# save 