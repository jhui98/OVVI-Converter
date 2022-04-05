# Jacob Hui - Clover Automation Functions

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

def get_departments(clover_wb):
    ws = clover_wb["Categories"]
    departments = []
    for cell in ws[1]:
        departments.append(cell.value)
    departments.pop(0)
    # print(departments)
    return departments 

def item_department_dict(clover_wb, departments):
    item_dict = {}
    ws = clover_wb["Categories"] # select categories ws
    for col in range(len(departments)):
        dep = str(departments[col]) # save department name
        # print(dep)
        for cell in ws[get_column_letter(col+2)]: # for every cell within column index
            if (cell.value != None) and (cell.value != dep): # cell not empty and not department
                item = str(cell.value)
                item_dict[item] = {"department": dep}
    return item_dict

def initialItemIstance(clover_wb, departments):
    items = []
    ws = clover_wb["Categories"] # select categories ws
    for col in range(len(departments)):
        dep = str(departments[col]) # save department name
        for cell in ws[get_column_letter(col+2)]: # for every cell within column index
            if (cell.value != None) and (cell.value != dep): # cell not empty and not department
                item = str(cell.value)
                item = Ovvi(item, dep)
                items.append(item)
    return items


def init_OVVI(folder):
    wb = Workbook() 
    dest_filename = folder + 'Item-PLU with Data.xlsx'   
    ws1 = wb.active
    ws1.title = "Item-PLU" 
    ws2 = wb.create_sheet(title="ModifierGroups")   
    
    
    wb.save(filename = dest_filename)

def findItem(itemsList, itemName):
    index = 0
    for item in itemsList:
        if item.itemName == itemName:
            print(item.itemName)
            print(item.itemDepartment)
            print(index)
    index += 1


class Ovvi:
    def __init__(self, itemName, department):
        self.itemName = itemName
        self.itemDepartment = department
        self.itemSellPrice = 0.00
        self.itemBarcode = None
        self.itemCost = 0.00
    # Item structure
        # item_name = {
        # department: "",
        # price: float(),
        # barcode: "",
        # cost: float(),
        # }

    def changeSellPrice(self, sellPrice): # change item sell price
        self.itemSellPrice = sellPrice

    def changeItemBarcode(self, barcode): # change item barcode
        self.itemBarcode = barcode

    def changeIemCost(self, itemCost): # change item cost
        self.itemCost = itemCost

    def leadingZeroBarcode(self): # change item barcode
        ogBarcode = self.itemBarcode 
        updatedBarcode = ogBarcode + "@0" + ogBarcode
        self.itemBarcode = updatedBarcode
    
