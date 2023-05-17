import pandas as pd
import os
from datetime import datetime


def get_most_recent_file(folder_path):
     #= "../Data Pool/DCT Process Results/Exported Result Files"

    # List all the files in the directory
    files = os.listdir(folder_path)

    # Filter out directories, leaving files only
    files = [f for f in files if os.path.isfile(os.path.join(folder_path, f))]

    # Get creation time of the files and sort them by it
    files.sort(key=lambda x: os.path.getctime(os.path.join(folder_path, x)))

    # Return the most recent file
    return files[-1]


def get_po_line_cost(po_number, project_number):
    folder_path = "../Data Pool/Ecosys API Data/PO Lines"

    # List all excel files in the directory
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    # Filter files that contain the project number
    matching_files = [f for f in excel_files if str(project_number) in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found."
        )

    # Get the most recent file
    most_recent_file = get_most_recent_file(folder_path)
    file_path = os.path.join(folder_path, most_recent_file)

    # Load the file into a DataFrame
    df = pd.read_excel(file_path)

    # Filter rows that match the PO number and calculate the total cost
    cost = df.loc[df['Cost Object ID'] == po_number, 'Cost Project Currency'].sum()

    return cost


def analyze_by_po_number(project_number):

    folder_path = "../Data Pool/DCT Process Results/Exported Result Files"
    most_recent_file = get_most_recent_file(folder_path)
    file_path = os.path.join(folder_path, most_recent_file)

    # Load the most recent file into a DataFrame
    df = pd.read_excel(file_path)

    # Replace the commas in the 'PO Number' column with a whitespace to separate multiple PO numbers
    df['PO Number'] = df['PO Number'].str.replace(',', ' ')

    # Create a new DataFrame which contains one row for each PO number
    expanded_df = df['PO Number'].str.split(' ', expand=True).stack().reset_index(level=-1, drop=True)
    expanded_df.name = 'PO Number'
    df_expanded = df.drop('PO Number', axis=1).join(expanded_df)

    # Calculate the cost for each PO number
    df_expanded['Cost'] = df_expanded['PO Number'].apply(get_po_line_cost, args=(project_number,))

    # Export to a new Excel file
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    output_folder_path = "../Data Pool/DCT Process Results/PO Analysis"
    os.makedirs(output_folder_path, exist_ok=True)
    output_file = os.path.join(output_folder_path, f"PO_Analysis_{project_number}_{timestamp}.xlsx")
    df_expanded.to_excel(output_file, index=False)

    print(f"Data exported to {output_file}")


# Test the function
project_number = "17033" # replace this with your actual project number
analyze_by_po_number(project_number)