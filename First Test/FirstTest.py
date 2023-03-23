import openpyxl

# Load the source workbook
source_wb = openpyxl.load_workbook("/Users/lopes/PycharmProjects/PythonAutomatization/First Test/Fichier 1.xlsx")

# Select the active worksheet in the source workbook
source_ws = source_wb.active

# Get the value in cell A5 of the source worksheet
value_to_copy = source_ws["A5"].value

# Load the destination workbook
dest_wb = openpyxl.load_workbook("/Users/lopes/PycharmProjects/PythonAutomatization/First Test/Fichier 2.xlsx")

# Select the active worksheet in the destination workbook
dest_ws = dest_wb.active

# Paste the value into cell B8 of the destination worksheet
dest_ws["B8"].value = value_to_copy

# Save the changes to the destination workbook
dest_wb.save("/Users/lopes/PycharmProjects/PythonAutomatization/First Test/Fichier 2.xlsx")
