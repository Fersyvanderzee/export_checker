from openpyxl import load_workbook
import sys
import re

workbook = load_workbook(filename="export.xlsx")
sheet = workbook.active

with open('result.txt', 'w') as file:

    message = " Programma wordt afgn esloten."

    skus = []
    additional_attributes = []

    #check
    if sheet["A1"].value != "sku":
        sys.exit("Kolom A is niet sku." + message)

    if sheet["Q1"].value != "additional_attributes":
        sys.exit("Kolom Q is niet additional_attributes." + message)

    for skus_iter in sheet["A"]:
        skus.append(skus_iter.value)

    for additional_attributes_iter in sheet["Q"]:
        additional_attributes.append(additional_attributes_iter.value)

    row = 0
    count = 0
    print('Producten die op tijdelijk aflopend staan:\n')
    for sku in skus:
        if additional_attributes[row] is not None and "saleartikel=Tijdelijk" in additional_attributes[row]:
            print(sku)
        row += 1
        count += 1

    print('\nDone')