# Jacob Hui - Clover Automation Functions

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

def get_departments(clover_wb):
    ws = clover_wb["Categories"]
    departments = []
    for cell in ws[1]:
        departments.append(cell.value)
    # departments.pop(0)
    # print(departments)
    return departments 

def item_department_dict(clover_wb, departments):
    items = {}
    ws = clover_wb["Categories"] # select categories ws
    for col in range(len(departments)):
        dep = str(departments[col]) # save department name
        # print(dep)
        for cell in ws[get_column_letter(col+2)]: # for every cell within column index
            if (cell.value != None) and (cell.value != dep): # cell not empty and not department
                item = str(cell.value)
                items[item] = {"department": dep}
                # print(cell.value)
    return items

def init_OVVI(folder):
    wb = Workbook() 
    dest_filename = folder + 'Item-PLU with Data.xlsx'   
    ws1 = wb.active
    ws1.title = "Item-PLU" 
    ws2 = wb.create_sheet(title="ModifierGroups")   
    
    
    wb.save(filename = dest_filename)
