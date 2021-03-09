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
    output = pd.DataFrame({'product_type': product_typecol,"product_websites": product_websitescol, "name": name_col,
                           "weight": weight_col,
                           "product_online": product_onlinecol, "tax_class_name": taxable_goodscol,
                           "visibility": visibility_col, "price": price_col, "url_key": url_keycol,
                           "meta_title": meta_titlecol, "meta_keywords": meta_keywordscol,
                           "meta_description": meta_descriptioncol})
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
    print(len(rows_list))
    print("price col complete")
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
    print("url key col complete")
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
    print(len(rows_list))
    print("meta_title col complete")
    return rows_list


def meta_keywords():
    rows = meta_title()
    rows.pop(0)
    rows.insert(0, "meta_keywords")
    print("meta_keywords col complete")
    return rows


def meta_description():
    rows = meta_title()
    rows.pop(0)
    rows.insert(0, "meta_description")
    print("meta_description col complete")
    return rows


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
