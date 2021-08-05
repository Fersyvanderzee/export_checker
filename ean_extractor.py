from openpyxl import load_workbook
import sys
import re

workbook = load_workbook(filename="export.xlsx")
sheet = workbook.active

with open('result.txt', 'w') as file:

    message = " Programma wordt afgesloten."

    skus = []
    additional_attributes = []

    #check
    if sheet["A1"].value != "sku":
        sys.exit("Kolom A is niet sku." + message)

    if sheet["Q1"].value != "additional_attributes":
        sys.exit("Kolom Q is niet additional_attributes." + message)

    #iteration
    for skus_iter in sheet["A"]:
        skus.append(skus_iter.value)

    for additional_attributes_iter in sheet["Q"]:
        additional_attributes.append(additional_attributes_iter.value)

    row = 0
    count = 0
    ean_count = 0
    for sku in skus:
        if additional_attributes[row] is not None and "ean" in additional_attributes[row]:
            found_ean = re.search("ean=(.+?),", additional_attributes[row])
            found_eenheid = re.search("eenheid=(.+?),", additional_attributes[row])
            if found_ean is not None and found_eenheid is not None:
                ean = found_ean.group()[4:-1]
                eenheid = found_eenheid.group()[8:-1]
                ean.strip()
                ean.strip('\t')
                eenheid.strip()
                eenheid.strip('\t')
                file.writelines(str(sku) + ',' + str(ean) + ',' + str(eenheid) +'\n')
                ean_count += 1
        row += 1
        count += 1


    print(f'Wrote {count} lines, {ean_count} lines containing an ean-code')
    ean_fil = ean_count / count * 100
    print(f'ean percentage: {ean_fil}%')