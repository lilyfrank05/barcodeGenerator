import random
import string
from openpyxl import Workbook


length = int(input("How long shall the barcode be? "))
composition = input("What will be in your barcode? \nNumbers only (1) \nLetters only (2) \nNumbers + Letters (3)"
                    "\nPlease enter 1, 2 or 3. ")
numberofBarcodes = int(input("How many barcodes do you need? "))
sheetPath = input("Where do you want to save it? E.g /Users/frank/Downloads\n")
barcode = ""
barcodes = []
i = 0

while i < numberofBarcodes:
    if composition == "1":
        for a in range(length):
            barcode += random.choice(string.digits)
    elif composition == "2":
        for b in range(length):
            barcode += random.choice(string.ascii_uppercase)
    elif composition == "3":
        for c in range(length):
            barcode += random.choice(string.digits + string.ascii_uppercase)
    if barcode not in barcodes:
        barcodes.append(barcode)
        i += 1
    barcode = ""
print(barcodes)

workbook = Workbook()
sheet = workbook.active
sheet.cell(row=1, column=1).value = "Barcode"
for l in range(len(barcodes)):
    sheet.cell(row=l+2, column=1).value = barcodes[l]
workbook.save(filename=sheetPath + "/barcodelist.xlsx")