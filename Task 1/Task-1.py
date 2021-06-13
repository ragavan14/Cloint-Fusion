import os
import shutil
import ClointFusion as cf

cf.OFF_semi_automatic_mode()

OLD_path = r"D:\clointfusion_projects\TASK-1\OriginalDocuments\Excel.xlsx"
NEW_path = r"D:\clointfusion_projects\TASK-1\CreatedDocuments\Excel.xlsx"
if not os.path.exists('CreatedDocuments'):
    os.makedirs('CreatedDocuments')
if not os.path.exists(NEW_path):
    shutil.copy(OLD_path, NEW_path)
# 1
header_columns = cf.excel_get_all_header_columns(NEW_path)
print(header_columns)

# 2
row_column_count = cf.excel_get_row_column_count(NEW_path)
row_count = row_column_count[0]
column_count = row_column_count[1]
print(row_column_count)

# 3
sheet_names = cf.excel_get_all_sheet_names(NEW_path)
print(sheet_names)

# 4
cf.excel_remove_duplicates(NEW_path)

# 5
cf.excel_sort_columns(NEW_path)

# 6

Dict = {"ID ": 1027, "OrderDate": "4/14/2020", "Region": "East", "Rep": "Jones", "Item": "Binder", "Units": 60,
        "UnitCost": 4.99, "Total": 499.1}

for i in range(0, column_count):
    cf.excel_set_single_cell(excel_path=NEW_path, columnName=header_columns[i], cellNumber=43,
                             setText=Dict[header_columns[i]])

# 7
cf.excel_split_the_file_on_row_count(excel_path=NEW_path, sheet_name="Split", rowSplitLimit=12,
                                     outputFolderPath=r"D:\clointfusion_projects\TASK-1\CreatedDocuments")

# 8
# Creating  dictionary
data = {1001: 95,
        1002: 50,
        1003: 36,
        1004: 27,
        1005: 56,
        1006: 60,
        1007: 75,
        1008: 90,
        1009: 32,
        1010: 60,
        1011: 90,
        1012: 29,
        1013: 81,
        1014: 35,
        1015: 2,
        1016: 16,
        1017: 28,
        1018: 64,
        1019: 15,
        1020: 96,
        1021: 67,
        1022: 74,
        1023: 46,
        1024: 87,
        1025: 4,
        1026: 7,
        1027: 50,
        1028: 66,
        1029: 96,
        1030: 53,
        1031: 80,
        1032: 5,
        1033: 62,
        1034: 55,
        1035: 42,
        1036: 3,
        1037: 7,
        1038: 76,
        1039: 57,
        1040: 14,
        1041: 11,
        1042: 94,
        1043: 28,

        }
