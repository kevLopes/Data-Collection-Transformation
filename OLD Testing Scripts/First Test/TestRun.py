import openpyxl
from openpyxl.chart import BarChart, Reference

# Create a workbook and select the active worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active

# Create some sample data
data = [
    ['Month', 'Sales'],
    ['Jan', 150],
    ['Feb', 200],
    ['Mar', 175],
    ['Apr', 225],
    ['May', 250],
    ['Jun', 300]
]

# Write the data to the worksheet
for row in data:
    worksheet.append(row)

# Create a bar chart
chart = BarChart()
chart.title = 'Sales by Month'
chart.x_axis.title = 'Month'
chart.y_axis.title = 'Sales'
data = Reference(worksheet, min_col=2, min_row=1, max_row=7, max_col=2)
labels = Reference(worksheet, min_col=1, min_row=2, max_row=7)
chart.add_data(data, titles_from_data=True)
chart.set_categories(labels)

# Add the chart to the worksheet
worksheet.add_chart(chart, 'A8')

# Save the workbook
workbook.save('sales.xlsx')
