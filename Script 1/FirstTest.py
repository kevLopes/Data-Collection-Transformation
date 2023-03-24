import openpyxl

# Open the first workbook and get the value from cell A5
wb1 = openpyxl.load_workbook(r'C:\Users\keven.deoliveiralope\Documents\Data Analyze Automatization\Scripts\Script 1\Fichie 1.xlsx')
ws1 = wb1.active
value_to_copy = ws1['A5'].value

# Open the second workbook and paste the value into cell B8
wb2 = openpyxl.load_workbook(r'C:\Users\keven.deoliveiralope\Documents\Data Analyze Automatization\Scripts\Script 1\Fichie 2.xlsx')
ws2 = wb2.active
ws2['B8'].value = value_to_copy

# Save the changes to the second workbook
wb2.save(r'C:\Users\keven.deoliveiralope\Documents\Data Analyze Automatization\Scripts\Script 1\Fichie 2.xlsx')