import backend
import sys
from openpyxl import load_workbook
from backend import *

# Path to the file to use, for example: export.xlsx
filename = 'export/export.xlsx'

# Filename for results, for example: result.txt
file = open('result.txt', 'w')

#Columns in Excel:
skus_excel = "A"
product_type_excel = "D"
categories_excel = "E"
weight_excel = "G"
product_online_excel = "H"
visibility_excel = "J"
additional_attributes_excel = "R"
qty_excel = "S"
use_config_backorders_excel = "X"
min_cart_qty_excel = "Y"
use_config_min_sale_qty_excel = "Z"
is_in_stock_excel = "AC"

try:
    workbook = load_workbook(filename=filename)
    sheet = workbook.active
    backend.log(f'Bestand met naam \'{filename}\' gevonden. Iteratie gestart.')
except FileNotFoundError:
    backend.log(f'Er is geen bestand met de naam \'{filename}\' gevonden.\n')
    sys.exit("Er is geen bestand met de naam 'export.xlsx' in de map 'export' gevonden." + backend.message)


######################################################################
# check and iteration
######################################################################

# check values
if sheet[skus_excel + "1"].value != "sku":
    sys.exit(f"Kolom {skus_excel} is niet sku. {message}")
if sheet[product_type_excel + "1"].value != "product_type":
    sys.exit(f"Kolom D {product_type_excel} is niet product_type. {message}")
if sheet[categories_excel + "1"].value != "categories":
    sys.exit(f"Kolom E {categories_excel} is niet categories. {message}")
if sheet[weight_excel + "1"].value != "weight":
    sys.exit(f"Kolom G {weight_excel} is niet weight. {message}")
if sheet[product_online_excel + "1"].value != "product_online":
    sys.exit(f"Kolom H {product_online_excel} is niet product_online. {message}")
if sheet[visibility_excel + "1"].value != "visibility":
    sys.exit(f"Kolom J {visibility_excel} is niet visibility. {message}")
if sheet[additional_attributes_excel + "1"].value != "additional_attributes":
    sys.exit(f"Kolom R {additional_attributes_excel} is niet additional_attributes. {message}")
if sheet[qty_excel + "1"].value != "qty":
    sys.exit(f"Kolom S {qty_excel} is niet qty. {message}")
if sheet[use_config_backorders_excel + "1"].value != "use_config_backorders":
    sys.exit(f"Kolom X {use_config_backorders_excel} is niet use_config_backorders. {message}")
if sheet[min_cart_qty_excel + "1"].value != "min_cart_qty":
    sys.exit(f"Kolom Y {min_cart_qty_excel} is niet min_cart_qty. {message}")
if sheet[use_config_min_sale_qty_excel + "1"].value != "use_config_min_sale_qty":
    sys.exit(f"Kolom Z {use_config_min_sale_qty_excel} is niet use_config_min_sale_qty. {message}")
if sheet[is_in_stock_excel + "1"].value != "is_in_stock":
    sys.exit(f"Kolom AC {is_in_stock_excel} is niet is_in_stock. {message}")


# iteration
skus = [skus_iter.value for skus_iter in sheet[skus_excel]]
product_type = [product_type_iter.value for product_type_iter in sheet[product_type_excel]]
categories = [categories_iter.value for categories_iter in sheet[categories_excel]]
weight = [weight_iter.value for weight_iter in sheet[weight_excel]]
product_online = [product_online_iter.value for product_online_iter in sheet[product_online_excel]]
visibility = [visibility_iter.value for visibility_iter in sheet[visibility_excel]]
additional_attributes = [additional_attributes_iter.value for additional_attributes_iter in sheet[additional_attributes_excel]]
qty = [qty_iter.value for qty_iter in sheet[qty_excel]]
use_config_backorders = [use_config_backorders_iter.value for use_config_backorders_iter in sheet[use_config_backorders_excel]]
min_cart_qty = [min_cart_qty_iter.value for min_cart_qty_iter in sheet[min_cart_qty_excel]]
use_config_min_sale_qty = [use_config_min_sale_qty_iter.value for use_config_min_sale_qty_iter in sheet[use_config_min_sale_qty_excel]]
is_in_stock = [is_in_stock_iter.value for is_in_stock_iter in sheet[is_in_stock_excel]]


######################################################################
# Tasks
######################################################################

# Aflopende artikelen controleren
name = 'Aflopende artikelen controleren'
print(name)
print("Deze artikelen staan op Op=Op of Tijdelijk aflopend en hebben geen voorraad meer:")
row = 0
count = 0
log_list = []
for sku in skus:
    if (additional_attributes[row] is not None) and ("stock_display=Yes" in additional_attributes[row]) \
            and (float(qty[row]) <= 0):
        if "saleartikel=Sale" in additional_attributes[row]:
            print(str(skus[row]) + ' - Sale')
            log_list.append(skus[row])
            count += 1
        elif "saleartikel=Tijdelijk" in additional_attributes[row]:
            print(str(skus[row]) + ' - Tijdelijk')
            log_list.append(skus[row])
            count += 1
        elif "saleartikel=out" in additional_attributes[row]:
            print(str(skus[row]) + ' - Out of stock')
            log_list.append(skus[row])
            count += 1
        else:
            print(str(skus[row]) + ' - checken')
            log_list.append(skus[row])
            count += 1
    row += 1
if count < 1:
    print("Klopt")
loglist(name, log_list)
print('\n')

# Alle op=op moet op visibility 'Catalog,search' staan
name = 'Alle op=op moet op visibility Catalog,search staan'
print(name)
print("Deze artikelen staat op Op=Op, maar niet op Catalog,search:")
row = 0
count = 0
log_list = []
for sku in skus:
    if additional_attributes[row] is not None and visibility[row] is not None:
        if visibility[row] != 'Catalog, Search' and "saleartikel=Op" in additional_attributes[row]:
            print(skus[row])
            log_list.append(skus[row])
            count += 1
    row += 1
if count < 1:
    print("Klopt")
loglist(name, log_list)
print('\n')


# Catalogus,zoeken + ingeschakeld + niet op voorraad
name = 'Catalogus,zoeken + ingeschakeld + niet op voorraad'
print(name)
row = 0
count = 0
log_list = []
for sku in skus:
    if visibility[row] == 'Catalog, Search' and product_online[row] == 1 and is_in_stock[row] == 0:
        if "saleartikel=out" not in additional_attributes[row] and float(qty[row]) > 0:
            print(skus[row])
            log_list.append(skus[row])
            count += 1
    row += 1
if count < 1:
    print("Klopt")
loglist(name, log_list)
print('\n')


# Config mag geen minimale afname hebben - Export
name = 'Config mag geen minimale afname hebben - Export'
print(name)
row = 0
count = 0
log_list = []
for sku in skus:
    if min_cart_qty[row] is not None:
        if product_type[row] == 'configurable' and min_cart_qty[row] != 10000:
            print(skus[row])
            log_list.append(skus[row])
            count += 1
    row += 1
if count < 1:
    print("Klopt")
loglist(name, log_list)
print('\n')


# Config min cart qty check ivm Bol - Export
name = 'Config min cart qty check ivm Bol - Export'
print(name)
row = 0
count = 0
log_list = []
for sku in skus:
    if use_config_min_sale_qty[row] == 1 and min_cart_qty[row] != 10000:
        print(skus[row])
        log_list.append(skus[row])
        count += 1
    row += 1
if count < 1:
    print("Klopt")
else:
    print('Klopt niet. Open de SKU\'s in Magento. Haal config eraf en sla op -> nu zie je een verkeerd aantal staan.')
loglist(name, log_list)
print('\n')


# Geen backorders en tonen actuele voorraad
name = 'Geen backorders en tonen actuele voorraad'
print(name)
row = 0
count = 0
checken_producten = []
log_list = []
for sku in skus:
    if use_config_backorders[row] == 0:
        if "stock_display=No" in additional_attributes[row]:
            checken_producten.append(skus[row])
            log_list.append(skus[row])
            count += 1
    row += 1
if count < 1:
    print("Klopt: geen producten op 'No backorders' en 'Toon actuele voorraadaantal: nee'")
else:
    print("Producten die verkeerd staan:")
    for sku in checken_producten:
        print(sku)
loglist(name, log_list)
print("\n")

name = 'Check of alle producten die op No backorders staan ook een sale artikel hebben'
log_list = []
print(name)
row = 0
backorder_count = 0
saleartikel_tijdelijk_count = 0
saleartikel_op_count = 0
saleartikel_sale_count = 0
saleartikel_out_count = 0

checken_saleartikel_producten = []
wrong_products = []

for sku in skus:
    if use_config_backorders[row] == 0:
        backorder_count += 1
        if "saleartikel=Tijdelijk" in additional_attributes[row]:
            checken_saleartikel_producten.append(skus[row])
            saleartikel_tijdelijk_count += 1
        elif "saleartikel=Op" in additional_attributes[row]:
            checken_saleartikel_producten.append(skus[row])
            saleartikel_op_count += 1
        elif "saleartikel=Sale" in additional_attributes[row]:
            checken_saleartikel_producten.append(skus[row])
            saleartikel_sale_count += 1
        elif "saleartikel=out" in additional_attributes[row]:
            checken_saleartikel_producten.append(skus[row])
            saleartikel_out_count += 1
        else:
            wrong_products.append(skus[row])

    row += 1

print("Producten op 'No backorders': " + str(backorder_count))
print("Tijdelijk aflopend: " + str(saleartikel_tijdelijk_count))
print("Op=Op: " + str(saleartikel_op_count))
print("Sale: " + str(saleartikel_sale_count))
print("Out of stock: " + str(saleartikel_out_count))
result_sale_artikel = saleartikel_tijdelijk_count + saleartikel_op_count \
                      + saleartikel_sale_count + saleartikel_out_count
print("Totaal: " + str(result_sale_artikel))
if result_sale_artikel == backorder_count:
    print("Klopt.")
else:
    print("Klopt niet. Check deze producten:")
    for sku in wrong_products:
        print(sku)
loglist(name, log_list)
print('\n')


# Uitgeschakeld mag niet op in stock
name = 'Uitgeschakeld mag niet op in stock'
print(name)
row = 0
count = 0
log_list = []
wabco = [103006001, 103006002, 103006003, 103006004,
         103006005, 103006006, 103006007, 103006008,
         103006009, 103006010, 103006011]
for sku in skus:
    if product_online[row] == 2 and is_in_stock[row] == 1:
        if sku in wabco:
            print(str(skus[row]) + " - Wabco")
        else:
            print(skus[row])
        log_list.append(skus[row])
        count += 1
    row += 1
if count < 1:
    print("Klopt")
loglist(name, log_list)
print('\n')


# Klantspecifiek niet in categorie
name = 'Klantspecifiek niet in categorie'
print(name)
row = 0
count = 0
log_list = []
for sku in skus:
    if additional_attributes[row] is not None and "moet_in_categorie=Nee" in additional_attributes[row] \
            and categories[row] is not None:
        print(skus[row])
        log_list.append(skus[row])
        count += 1
    row += 1

if count < 1:
    print("Klopt.")
else:
    print("Klopt niet.")
loglist(name, log_list)
print("\n")

# count lines
line_count = 0
for sku in skus:
    line_count += 1

print(f'Done. {line_count} lines processed.')
log(f'Done. {line_count} lines processed.')
