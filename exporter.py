from openpyxl import load_workbook
import sys
import re

"""
4-8-2021 FZE: In het export-bestand staan 3 rijen die er niet in horen. Deze hebben een lege waar de in
additional attributes. Dat maakt het verschil. Het programma processed 1 line meer dan dat er producten zijn
omdat de headers ook worden meegenomen.
"""

# 5-8-2021 FZE: Let op: bij volgende export uit ODBC 'aflopend' en 'cdinactief' meenemen.

workbook = load_workbook(filename="export.xlsx")
sheet = workbook.active

with open('result_exporter.txt', 'w') as file:

    message = " Programma wordt afgesloten."

    skus = []
    visibility = []
    additional_attributes = []
    product_online = []
    is_in_stock = []
    land = []

    #check
    if sheet["A1"].value != "sku":
        sys.exit("Kolom A is niet sku." + message)

    if sheet["B1"].value != "store_view_code":
        sys.exit("Kolom B is niet store_view_code." + message)

    if sheet["I1"].value != "visibility":
        sys.exit("Kolom I is niet visibilty." + message)

    if sheet["Q1"].value != "additional_attributes":
        sys.exit("Kolom Q is niet additional_attributes." + message)

    if sheet["G1"].value != "product_online":
        sys.exit("Kolom G is niet product_online." + message)

    if sheet["AB1"].value != "is_in_stock":
        sys.exit("Kolom AB is niet is_in_stock." + message)

    #iteration
    for skus_iter in sheet["A"]:
        skus.append(skus_iter.value)

    for land_iter in sheet["B"]:
        land.append(land_iter.value)

    for product_online_iter in sheet["G"]:
        product_online.append(product_online_iter.value)

    for additional_attributes_iter in sheet["Q"]:
        additional_attributes.append(additional_attributes_iter.value)

    for is_in_stock_iter in sheet["AB"]:
        is_in_stock.append(is_in_stock_iter.value)

    for visibility_iter in sheet["I"]:
        visibility.append(visibility_iter.value)

    row = 0
    count = 0
    for sku in skus:
        if additional_attributes[row] is not None and land[row] is None:
            found_eenheid = re.search("eenheid=(.+?),", additional_attributes[row])
            found_saleartikel = re.search("saleartikel=(.+?),", additional_attributes[row])
            if found_eenheid is not None:
                eenheid = found_eenheid.group()[8:-1]
            if found_saleartikel is not None:
                saleartikel = found_saleartikel.group()[12:-1]
            eenheid.strip()
            eenheid.strip('\t')
            zoekcode = str(sku + eenheid)
            file.writelines(str(sku) + ',' + str(eenheid if found_eenheid is not None else 'geen') + ','
                            + str('ja' if product_online[row] == 1 else 'nee') + ','
                            + str('ja' if is_in_stock[row] == 1 else 'nee') + ','
                            + (str('"' + visibility[row] + '"' if visibility[row] is not None else 'geen')) + ','
                            + str(saleartikel if found_saleartikel is not None else 'geen')
                            + str(zoekcode if found_eenheid is not None else 'geen') +
                            '\n')
            count += 1
        row += 1

    print(f'processed {count} lines')
