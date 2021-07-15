from openpyxl import load_workbook

workbook = load_workbook(filename="export.xlsx")
sheet = workbook.active


# iterate all values to arrays
skus = []
product_online = []
is_in_stock = []
product_type = []
min_cart_qty = []
use_config_min_sale_qty = []
additional_attributes = []
categories = []
visibility = []
qty = []
use_config_backorders = []


# iterating
for skus_iter in sheet["A"]:
    skus.append(skus_iter.value)

for product_online_iter in sheet["G"]:
    product_online.append(product_online_iter.value)

for is_in_stock_iter in sheet["AD"]:
    is_in_stock.append(is_in_stock_iter.value)

for product_type_iter in sheet["D"]:
    product_type.append(product_type_iter.value)

for min_cart_qty_iter in sheet["Z"]:
    min_cart_qty.append(min_cart_qty_iter.value)

for use_config_min_sale_qty_iter in sheet["AA"]:
    use_config_min_sale_qty.append(use_config_min_sale_qty_iter.value)

for additional_attributes_iter in sheet["S"]:
    additional_attributes.append(additional_attributes_iter.value)

for categories_iter in sheet["E"]:
    categories.append(categories_iter.value)

for visibility_iter in sheet["I"]:
    visibility.append(visibility_iter.value)

for qty_iter in sheet["T"]:
    qty.append(qty_iter.value)

for use_config_backorders_iter in sheet["Y"]:
    use_config_backorders.append(use_config_backorders_iter.value)


# Aflopende artikelen controleren
print("Aflopende artikelen controleren:")
print("Deze artikelen staan op Op=Op of Tijdelijk aflopend en hebben geen voorraad meer:")
row = 0
count = 0
for sku in skus:
    if additional_attributes[row] is not None and "stock_display=Yes" in additional_attributes[row] \
            and float(qty[row]) <= 0:
        if "saleartikel=out" not in additional_attributes[row]:
            print(skus[row])
            count += 1
    row += 1
if count < 1:
    print("Klopt")
print('\n')


# Alle op=op moet op visibility 'Catalog,search' staan
print("Alle op=op moet op visibility 'Catalog,search' staan")
print("Deze artikelen staat op Op=Op, maar niet op Catalog,search:")
row = 0
count = 0
for sku in skus:
    if additional_attributes[row] is not None and visibility[row] is not None:
        if visibility[row] != 'Catalog, Search' and "saleartikel=Op" in additional_attributes[row]:
            print(skus[row])
            count += 1
    row += 1
if count < 1:
    print("Klopt")
print('\n')


# Catalogus,zoeken + ingeschakeld + niet op voorraad
print('Catalogus,zoeken + ingeschakeld + niet op voorraad:')
row = 0
count = 0
for sku in skus:
    if visibility[row] == 'Catalog, Search' and product_online[row] == 1 and is_in_stock[row] == 0:
        if float(qty[row]) > 0:
            print(skus[row])
            count += 1
    row += 1
if count < 1:
    print("Klopt")
print('\n')


# Config mag geen minimale afname hebben - Export
print('Config mag geen minimale afname hebben - Export:')
row = 0
count = 0
for sku in skus:
    if min_cart_qty[row] is not None:
        if product_type[row] == 'configurable' and min_cart_qty[row] != 10000:
            print(skus[row])
            count += 1
    row += 1
if count < 1:
    print("Klopt")
print('\n')


# Config min cart qty check ivm Bol - Export
print('Config min cart qty check ivm Bol - Export:')
row = 0
count = 0
for sku in skus:
    if use_config_min_sale_qty[row] == 1 and min_cart_qty[row] != 10000:
        print(skus[row])
        count += 1
    row += 1
if count < 1:
    print("Klopt")
print('\n')


# Geen backorders en tonen actuele voorraad
print("Geen backorders en tonen actuele voorraad:")
row = 0
count = 0
checken_producten = []
for sku in skus:
    if use_config_backorders[row] == 0:
        if "stock_display=No" in additional_attributes[row]:
            checken_producten.append(skus[row])
            count += 1
    row += 1
if count < 1:
    print("Klopt: geen producten op 'No backorders' en 'Toon actuele voorraadaantal: nee'")
else:
    print("Producten die verkeerd staan:")
    for sku in checken_producten:
        print(sku)
print("\n")

print("Check of alle producten die op 'No backorders' staan ook een sale artikel hebben:")
row = 0
backorder_count = 0
saleartikel_tijdelijk_count = 0
saleartikel_op_count = 0
saleartikel_sale_count = 0
saleartikel_out_count = 0

checken_saleartikel_producten = []

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
    print("Klopt niet:")
    for sku in checken_saleartikel_producten:
        print(sku)

print('\n')


# Uitgeschakeld mag niet op in stock
print('Uitgeschakeld mag niet op in stock:')
row = 0
count = 0
for sku in skus:
    if product_online[row] == 2 and is_in_stock[row] == 1:
        print(skus[row])
        count += 1
    row += 1
if count < 1:
    print("Klopt")
print('\n')


# Klantspecifiek niet in categorie
print("Klantspecifiek niet in categorie")
row = 0
count = 0
for sku in skus:
    if additional_attributes[row] is not None and "moet_in_categorie=Nee" in additional_attributes[row] \
            and categories[row] is not None:
        print(skus[row])
        count += 1
    row += 1

if count < 1:
    print("Klopt.")
else:
    print("Klopt niet.")

print("\n")
print('Check klaar.')
