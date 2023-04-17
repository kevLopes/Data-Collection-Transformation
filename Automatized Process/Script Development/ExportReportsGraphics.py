import os
import openpyxl
from openpyxl.chart import ScatterChart, Reference, Series
from datetime import datetime


def data_graphic_reports_func():
    # Ask the user for the project ID
    project_id = input("Please enter the Project ID: ")

    # Find the Materials Data Organized folder
    folder_path = os.path.join(os.getcwd(), "Materials Data Organized")
    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' not found.")
        exit()

    # Find all Excel files with the project ID in the name
    files = [f for f in os.listdir(folder_path) if project_id in f and f.endswith('.xlsx')]

    # Define the columns to analyze
    columns = ["Tag Number", "ID", "Project Number", "Product Code", "Commodity Code",
               "Service Description", "Pipe Base Material", "Material", "LineNumber",
               "SBM scope", "Total QTY to commit", "Quantity UOM", "Unit Weight",
               "Unit Weight UOM", "Total NET weight", "SIZE"]

    # Create a dictionary to store the data
    data_dict = {}
    for column in columns:
        data_dict[column] = []

    # Loop through each file and extract the data
    for file in files:
        # Open the file
        try:
            workbook = openpyxl.load_workbook(os.path.join(folder_path, file))
        except Exception as e:
            print(f"Error: Could not load workbook {file} - {e}")
            continue

        # Loop through each worksheet
        for worksheet in workbook:
            # Loop through each row in the worksheet
            for row in worksheet.iter_rows(min_row=2, values_only=True):
                # Loop through each column in the row
                for i, column in enumerate(columns):
                    data_dict[column].append(row[i])

    # Create a new workbook to store the results
    result_workbook = openpyxl.Workbook()

    # Create a worksheet to store the data
    data_worksheet = result_workbook.create_sheet(title="Data")

    # Write the data to the worksheet
    for i, column in enumerate(columns):
        data_worksheet.cell(row=1, column=i + 1, value=column)
        for j, value in enumerate(data_dict[column]):
            data_worksheet.cell(row=j + 2, column=i + 1, value=value)

    # Create a scatter chart for each file
    chart = ScatterChart()
    chart.title = f"{project_id} - Total NET weight vs Unit Weight"
    chart.x_axis.title = 'Total NET weight'
    chart.y_axis.title = 'Unit Weight'
    chart.legend = None

    for file in files:
        # Create a new series for the file
        series = Series(
            Reference(data_worksheet, min_col=columns.index("Total NET weight") + 1, min_row=2,
                      max_row=len(data_dict[columns[0]])),
            Reference(data_worksheet, min_col=columns.index("Unit Weight") + 1, min_row=2,
                      max_row=len(data_dict[columns[0]])),
            title=file
        )

        # Add the series to the chart
        chart.series.append(series)

    # Add the chart to the worksheet
    chart_worksheet = result_workbook.create_sheet(title="Chart")
    chart_worksheet.add_chart(chart, "A1")

    # Get the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    # Save the result workbook
    result_filename = f"Result Graphic {project_id}_{timestamp}.xlsx"
    result_workbook.save(os.path.join(folder_path, result_filename))

    print("Done!")
