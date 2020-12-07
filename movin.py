# Reading an excel file using Python
import xlrd
import sqlite3
conn = sqlite3.connect('inventory.db')
# file to parse
file = ('inv.xlsx')
loc = file



# To open Workbook
wb = xlrd.open_workbook(loc)
#declaring which index of the sheet is read IE the bottom info
sheet = wb.sheet_by_index(5)

class Inv:
    def __init__(self, equip, asset, serial, manu, model, fname, lname, supplier, ponum):
        self.equip = equip
        self.asset = asset
        self.serial = serial
        self.manu = manu
        self.model = model
        self.fname = fname
        self.lname = lname
        self.supplier = supplier
        self.ponum = ponum


c = conn.cursor()

# Create table
c.execute('''CREATE TABLE inventory
             (Equipment text, AssestCode text, SerialNumber text, Manufacute text, Model text, FirstName text, LastName text, Supplier text, PONuber text)''')

# Insert a row of data

# Save (commit) the changes

# We can also close the connection if we are done with it.
# Just be sure any changes have been committed or they will be lost.

# For row 0
#  and column 0
counter = 0
for i in range(sheet.nrows):
    counter += 1
    # easy work around for replacing extra chars in the data
    firstName = sheet.cell_value(i, 5).replace("'", '')
    lastName = sheet.cell_value(i, 6).replace("'", '')
# creating the Inventory class object
    ok = Inv(sheet.cell_value(i, 0),
             sheet.cell_value(i, 1), 
             sheet.cell_value(i, 2), 
             sheet.cell_value(i, 3), 
             sheet.cell_value(i, 4), 
             firstName, 
             lastName, 
             sheet.cell_value(i, 7), 
             sheet.cell_value(i, 8))
    # c.execute(f"INSERT INTO inventory VALUES ('{ok.equip}', '{ok.asset}', '{ok.serial}', '{ok.manu}', '{ok.model}', '{ok.fname}', '{ok.lname}', '{ok.supplier}', '{ok.ponum}')")
    # executing the class object in the loop , setting the value to match the feilds and then send objs to the DB 
    c.execute(f"INSERT INTO inventory VALUES ('{ok.equip}', '{ok.asset}', '{ok.serial}', '{ok.manu}', '{ok.model}', '{ok.fname}', '{ok.lname}', '{ok.supplier}', '{ok.ponum}')")

    # c.execute(f"INSERT INTO stocks VALUES ('2006-01-{counter}','BUY{counter}','RHAT{counter}',100,{counter}.14)")

conn.commit()

# invSheet = open('inventory.txt', 'w')
# invSheet.write(invArr)
conn.close()

# print(invArr)
