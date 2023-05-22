import pandas as pd
import os
from datetime import datetime
import numpy as np



def get_most_recent_file(folder_path, matching_files):
    latest_time = None
    latest_file = None

    for file in matching_files:
        file_path = os.path.join(folder_path, file)
        file_time = os.path.getmtime(file_path)
        if latest_time is None or file_time > latest_time:
            latest_time = file_time
            latest_file = file

    return latest_file


def extract_po_numbers_and_costs(project_number, keyword):
    folder_path = "../Data Pool/DCT Process Results/Exported Result Files"
    new_folder_path = "../Data Pool/Ecosys API Data/PO Lines"

    excel_files = [
        f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    matching_files = [
        f for f in excel_files if str(project_number) in f and keyword in f and "Tag" in f
    ]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{keyword}' were found."
        )

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    # Check if 'PO Number' column exists
    if 'PO Number' not in df.columns:
        raise ValueError("PO Number column not found in the Excel sheet.")

    # Extract unique PO Numbers
    po_numbers = set()
    for val in df['PO Number'].dropna().values:
        # Split values by comma and strip white spaces
        for po_number in str(val).split(','):
            po_numbers.add(po_number.strip())

    print(f"Found {len(po_numbers)} unique PO numbers.")

    # Get the most recent file in the new folder
    new_excel_files = [
        f for f in os.listdir(new_folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    new_matching_files = [
        f for f in new_excel_files if str(project_number) in f
    ]

    if not new_matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found in {new_folder_path}."
        )

    most_recent_new_file = get_most_recent_file(new_folder_path, new_matching_files)
    new_file_path = os.path.join(new_folder_path, most_recent_new_file)
    new_df = pd.read_excel(new_file_path)

    # Check if 'Cost Project Currency' and 'Tag Number' columns exist in the new DataFrame
    if 'Cost Project Currency' not in new_df.columns or 'Cost Object ID' not in new_df.columns or 'Tag Number' not in new_df.columns:
        raise ValueError("Required columns not found in the Excel sheet.")

    # Sum up the cost per PO Number and check for empty Tag Numbers
    cost_data = []
    for po_number in po_numbers:
        po_rows = new_df[new_df['Cost Object ID'] == po_number]
        total_cost = po_rows['Cost Project Currency'].sum()
        remark = "PO Lines related to PO found with empty Tag number" if po_rows['Tag Number'].isna().any() else np.nan
        cost_data.append({
            "Project Number": project_number,
            "PO Number": po_number,
            "Total Cost": total_cost,
            "Currency": "USD",
            "Remarks": remark
        })

    # Convert to DataFrame and save to a new Excel file
    cost_df = pd.DataFrame(cost_data)
    output_folder_path = os.path.join(folder_path, "PO Analyze")
    os.makedirs(output_folder_path, exist_ok=True)  # Create the folder if it doesn't exist
    cost_df.to_excel(os.path.join(output_folder_path, "Cost_related_to_Piping_PO.xlsx"), index=False)




extract_po_numbers_and_costs("17033", "Piping")