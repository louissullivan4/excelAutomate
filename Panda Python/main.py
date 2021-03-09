import pandas as pd


def main():
    write_to_excel()


def write_to_excel():
    writer = pd.ExcelWriter('output.xlsx', engine="openpyxl")
    product_typecol = product_type()
    name_col = name()
    weight_col = weight()
    product_onlinecol = product_online()
    output = pd.DataFrame({'product_type': product_typecol, "name": name_col, "weight": weight_col,
                           "product_online": product_onlinecol})
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
            group = row.item()
        else:
            rows_list.append("simple")
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
