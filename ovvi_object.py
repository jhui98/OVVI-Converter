# Ovvi object structure

class Ovvi:
    def __init__(self, itemName, department):
        self.itemDepartment = department
        self.itemNumber = ""
        self.itemName = itemName
        self.modifierGroups = ""
        self.description = ""
        self.itemBarcode = ""
        self.itemCost = "0"
        self.itemSellPrice = "0"
        self.inStock = "0"
        self.tax1 = "TRUE"
        self.displayInMenu = "TRUE"
        self.isInventoryItem = "TRUE"
        self.isFoodStampable = "FALSE"
        self.beverageDeposit = ""

    # Item structure
        # item_name = {
        # department: "",
        # price: float(),
        # barcode: "",
        # cost: float(),
        # }
    
    # def change(self, ): # change 
    #     self. = 

    def changeitemDepartment(self, itemDepartment): # change itemDepartment
        self.itemDepartment = itemDepartment

    def changeitemName(self, itemName): # change itemName
        self.itemName = itemName

    def changeDescription(self, description): # change description
        self.description = description

    def changeItemBarcode(self, barcode): # change item barcode
        self.itemBarcode = barcode
    
    def changeIemCost(self, itemCost): # change item cost
        if itemCost != "":
            self.itemCost = itemCost

    def changeSellPrice(self, sellPrice): # change item sell price
        if sellPrice != "":
            self.itemSellPrice = sellPrice

    def changeinStock(self, inStock): # change item inStock
        self.inStock = inStock





    
