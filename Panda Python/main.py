import pandas as pd


def main():
    write_to_excel()


def write_to_excel():
    writer = pd.ExcelWriter('output.xlsx', engine="openpyxl")
    product_typecol = product_type()
    product_websitescol = product_websites()
    name_col = name()
    weight_col = weight()
    product_onlinecol = product_online()
    taxable_goodscol = tax_class_name()
    visibility_col = visibility()
    price_col = price()
    url_keycol = url_key()
    meta_titlecol = meta_title()
    meta_keywordscol = meta_keywords()
    meta_descriptioncol = meta_description()
    additional_attributescol = additional_attributes()
    qty_col = qty()
    is_in_stockcol = is_in_stock()
    website_idcol = website_id()
    outdict = {'product_type': product_typecol, "product_websites": product_websitescol, "name": name_col,
               "weight": weight_col,
               "product_online": product_onlinecol, "tax_class_name": taxable_goodscol,
               "visibility": visibility_col, "price": price_col, "url_key": url_keycol,
               "meta_title": meta_titlecol, "meta_keywords": meta_keywordscol,
               "meta_description": meta_descriptioncol, "additional_attributes": additional_attributescol,
               "qty": qty_col, "is_in_stock": is_in_stockcol, "website_id": website_idcol,
               "related_skus": ["related_skus"],
               "": ["related_position"], "crosssell_skus": ["crosssell_skus"],
               "crosssell_position": ["crosssell_position"], "upsell_skus": ["upsell_skus"],
               "upsell_position": ["upsell_position"]}
    output = pd.DataFrame.from_dict(outdict, orient='index')
    output = output.transpose()
    output.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
    writer.save()


def product_type():
    names = pd.read_excel('input.xlsx', sheet_name="new codes",
                          usecols=['WEB DESCRIPTION'])
    group = names.loc[0, 'WEB DESCRIPTION']
    rows_list = ["product_type"]
    for index, row in names.iterrows():
        if row.item() != group:
            rows_list.append("configurable")
            rows_list.append("simple")
            group = row.item()
        else:
            rows_list.append("simple")
    rows_list.append("configurable")
    print(len(rows_list))
    return rows_list


def product_websites():
    rows = name()
    rows.pop(0)
    numrows = len(rows)
    rows_list = ["base"]
    for i in range(numrows):
        rows_list.append("base")
    print(len(rows_list))
    return rows_list


def name():
    names = pd.read_excel('input.xlsx', sheet_name="new codes",
                          usecols=['WEB DESCRIPTION'])
    colours = pd.read_excel('input.xlsx', sheet_name="new codes",
                            usecols=["COLOUR"])
    sizes = pd.read_excel('input.xlsx', sheet_name="new codes",
                          usecols=["SIZE"])
    group = names.loc[0, 'WEB DESCRIPTION']
    rows_list = ["name"]
    for index, row in names.iterrows():
        if row.item() != group:
            rows_list.append(group)
            val = row.item()
            new = val + " " + colours.loc[index, 'COLOUR'] + " " + sizes.loc[index, 'SIZE']
            rows_list.append(new)
            group = row.item()
        else:
            val = row.item()
            new = val + " " + colours.loc[index, 'COLOUR'] + " " + sizes.loc[index, 'SIZE']
            rows_list.append(new)
    final = group
    rows_list.append(final)
    print(len(rows_list))
    return rows_list


def weight():
    rows = name()
    rows.pop(0)
    numrows = len(rows)
    rows_list = ["weight"]
    for i in range(numrows):
        rows_list.append(1)
    print(len(rows_list))
    return rows_list


def product_online():
    rows = name()
    rows.pop(0)
    numrows = len(rows)
    rows_list = ["product_online"]
    for i in range(numrows):
        rows_list.append(1)
    print(len(rows_list))
    return rows_list


def tax_class_name():
    rows = name()
    rows.pop(0)
    numrows = len(rows)
    rows_list = ["product_online"]
    for i in range(numrows):
        rows_list.append("Taxable Goods")
    print(len(rows_list))
    return rows_list


def visibility():
    names = pd.read_excel('input.xlsx', sheet_name="new codes",
                          usecols=['WEB DESCRIPTION'])
    group = names.loc[0, 'WEB DESCRIPTION']
    rows_list = ["visibility"]
    for index, row in names.iterrows():
        if row.item() != group:
            rows_list.append("Catalog, Search")
            rows_list.append("Not Visible Individually")
            group = row.item()
        else:
            rows_list.append("Not Visible Individually")
    rows_list.append("Catalog, Search")
    print(len(rows_list))
    return rows_list


def price():
    names = pd.read_excel('input.xlsx', sheet_name="new codes",
                          usecols=['WEB DESCRIPTION'])
    prices = pd.read_excel('input.xlsx', sheet_name="new codes",
                           usecols=["PRICE"])
    group = names.loc[0, 'WEB DESCRIPTION']
    current_price = prices.loc[0, 'PRICE']
    rows_list = ["price"]
    for index, row in names.iterrows():
        if row.item() != group:
            rows_list.append(current_price)
            group = row.item()
            current_price = prices.loc[index, 'PRICE']
            rows_list.append(current_price)
        else:
            new = prices.loc[index, 'PRICE']
            rows_list.append(new)
    final = current_price
    rows_list.append(final)
    print(len(rows_list))
    return rows_list


def url_key():
    rows = name()
    rows.pop(0)
    rows_list = ["url_key"]
    for val in rows:
        val = val.lower()
        val = val.replace(" ", "-").replace("'", "").replace("(", "").replace(")", "").replace("/", "-")
        rows_list.append(val)
    print(len(rows_list))
    return rows_list


def meta_title():
    names = pd.read_excel('input.xlsx', sheet_name="new codes",
                          usecols=['WEB DESCRIPTION'])
    dept = pd.read_excel('input.xlsx', sheet_name="new codes",
                         usecols=['DEPARTMENT'])
    rows_list = ["meta_title"]
    group = names.loc[0, 'WEB DESCRIPTION']
    current_dept = dept.loc[0, 'DEPARTMENT']
    for index, row in names.iterrows():
        if row.item() != group:
            rows_list.append(current_dept)
            group = row.item()
            current_department = dept.loc[index, 'DEPARTMENT']
            rows_list.append(current_department)
        else:
            new = dept.loc[index, 'DEPARTMENT']
            rows_list.append(new)
    final = current_dept
    rows_list.append(final)
    print(len(rows_list))
    return rows_list


def meta_keywords():
    rows = meta_title()
    rows.pop(0)
    rows.insert(0, "meta_keywords")
    return rows


def meta_description():
    rows = meta_title()
    rows.pop(0)
    rows.insert(0, "meta_description")
    return rows


def additional_attributes():
    names = pd.read_excel('input.xlsx', sheet_name="new codes",
                          usecols=['ITEM DESCRIPTION'])
    colours = pd.read_excel('input.xlsx', sheet_name="new codes",
                            usecols=["COLOUR"])
    sizes = pd.read_excel('input.xlsx', sheet_name="new codes",
                          usecols=["SIZE"])
    brands = pd.read_excel('input.xlsx', sheet_name="new codes",
                           usecols=['DEPARTMENT'])
    activity = pd.read_excel('input.xlsx', sheet_name="new codes",
                             usecols=['ACTIVITY'])
    group = names.loc[0, 'ITEM DESCRIPTION']
    rows_list = ["additional_attributes"]
    mens = "M"
    womens = "W"
    kids = ["Y", "G", "B"]
    size_guide = ",size="
    for index, row in names.iterrows():
        first_char = row[0][0]
        if first_char == womens:
            size_guide = ",womens_size="
        elif first_char in kids:
            size_guide = ",kids_size="
        elif first_char != womens and first_char == mens and first_char not in kids:
            size_guide = ",size="
        if row.item() != group:
            current_row = "colour=" + colours.loc[index, 'COLOUR'] + size_guide + sizes.loc[
                index, 'SIZE'] + ",brands=" + \
                          brands.loc[index, 'DEPARTMENT'] + ",activity=" + activity.loc[index, 'ACTIVITY']
            rows_list.append(None)
            group = row.item()
            rows_list.append(current_row)
        else:
            new = "colour=" + colours.loc[index, 'COLOUR'] + size_guide + sizes.loc[index, 'SIZE'] + ",brands=" + \
                  brands.loc[index, 'DEPARTMENT'] + ",activity=" + activity.loc[index, 'ACTIVITY']
            rows_list.append(new)
    rows_list.append(None)
    print(len(rows_list))
    return rows_list


def qty():
    names = pd.read_excel('input.xlsx', sheet_name="new codes",
                          usecols=['WEB DESCRIPTION'])
    group = names.loc[0, 'WEB DESCRIPTION']
    rows_list = ["gty"]
    for index, row in names.iterrows():
        if row.item() != group:
            rows_list.append(0)
            rows_list.append(1)
            group = row.item()
        else:
            rows_list.append(1)
    rows_list.append(0)
    print(len(rows_list))
    return rows_list


def is_in_stock():
    rows = name()
    rows.pop(0)
    numrows = len(rows)
    rows_list = ["is_in_stock"]
    for i in range(numrows):
        rows_list.append(1)
    print(len(rows_list))
    return rows_list


def website_id():
    rows = name()
    rows.pop(0)
    numrows = len(rows)
    rows_list = ["website_id"]
    for i in range(numrows):
        rows_list.append(0)
    print(len(rows_list))
    return rows_list



if __name__ == '__main__':
    main()

# def input_files():
#     while True:
#         try:
#             input_file = input("What file do you want to work on?: ")
#             output_file = input("where do you want to save the file?: ")
#         except FileNotFoundError:
#             print("Sorry this is not a valid file. Please enter one that exists in this folder")
#             continue
#         else:
#             return input_file, output_file
#
# def out_put():
#     while True:
#         try:
#             input_file = input("What file do you want to work on?: ")
#             output_file = input("where do you want to save the file?: ")
#         except FileNotFoundError:
#             print("Sorry this is not a valid file. Please enter one that exists in this folder")
#             continue
#         else:
#             return input_file, output_file
