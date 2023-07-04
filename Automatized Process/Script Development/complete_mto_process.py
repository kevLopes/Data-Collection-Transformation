import os
import pandas as pd
from datetime import datetime

folder_path = "../Data Pool/Data Hub Materials"


#Complete analyze of those different MTO files - Material type, Weight, Quantities
def complete_mto_data_analyze(project_number):

    # Call each function and collect their data
    piping_data = get_piping_mto_data(project_number)
    valve_data = get_valve_mto_data(project_number)
    bolt_data = get_bolt_mto_data(project_number)
    bent_data = get_bent_mto_data(project_number)
    structure_data = get_structure_mto_data(project_number)
    specialpip_data = get_specialpip_mto_data(project_number)

    # Combine the data into a data frame
    data_frame = pd.DataFrame({
        "Piping Data": piping_data,
        "Valve Data": valve_data,
        "Bolt Data": bolt_data,
        "Bent Data": bent_data,
        "Structure Data": structure_data,
        "Special Piping Data": specialpip_data
    })

    # Pass the data frame to export functions
    export_complete_mto_excel(data_frame)
    export_complete_mto_graphics(data_frame)
    export_complete_mto_pdf(data_frame)


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


def get_piping_mto_data(project_number):
    material_type = "Piping"

    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if material_type in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    # Set the columns to extract
    columns_to_extract = ["Project Number", "Pipe Base Material", "SBM scope", "Total QTY to commit", "Quantity UOM",
                          "Unit Weight", "Unit Weight UOM", "Total NET weight"]

    # Check if the project number matches the specified one
    if df["Project Number"].astype(str).str.contains(str(project_number)).any():
        # Find the specified columns
        extract_columns = []
        for column in df.columns:
            if any(col in str(column) for col in columns_to_extract):
                extract_columns.append(column)

        # Filter the DataFrame and extract the desired data
        filtered_df = df[extract_columns]

        # Filter based on SBM scope and notnull values
        extract_df_sbm = filtered_df[(filtered_df['SBM scope'] == True) & (filtered_df['SBM scope'].notnull())]
        extract_df_yard = filtered_df[(filtered_df['SBM scope'] == False) & (filtered_df['SBM scope'].notnull())]

        piping_data = filtered_df.groupby(["Pipe Base Material", "Quantity UOM"]).agg({
            "Total QTY to commit": "sum",
            "Total NET weight": "sum",
            "Unit Weight": "mean"
        }).reset_index()

        piping_sbm_data = extract_df_sbm.groupby(["Pipe Base Material", "Quantity UOM"]).agg({
            "Total QTY to commit": "sum",
            "Total NET weight": "sum",
            "Unit Weight": "mean"
        }).reset_index()

        piping_data_yard = extract_df_yard.groupby(["Pipe Base Material", "Quantity UOM"]).agg({
            "Total QTY to commit": "sum",
            "Total NET weight": "sum",
            "Unit Weight": "mean"
        }).reset_index()

        return piping_data, piping_sbm_data, piping_data_yard

    else:
        raise ValueError(f"No data found for project number '{project_number}'.")



def get_valve_mto_data(project_number):
    material_type = "Valve"

    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if material_type in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    # Set the columns to extract
    columns_to_extract = ["Status", "Project Number", "General Material Description", "Quantity", "Weight", "Remarks"]

    # Check if the project number matches the specified one
    if df["Project Number"].astype(str).str.contains(str(project_number)).any():
        # Find the specified columns
        extract_columns = []
        for column in df.columns:
            if any(col in str(column) for col in columns_to_extract):
                extract_columns.append(column)

        # Filter the DataFrame and extract the desired data
        filtered_df = df[extract_columns]

        # Filter based on SBM scope and notnull values
        extract_df_yard = filtered_df[(filtered_df['Status'] == "OUT OF SCOPE") | (filtered_df['Remarks'].str.contains("Yard", na=False))]
        extract_df_sbm = filtered_df[(filtered_df['Status'] == "CONFIRMED") & (~filtered_df['Remarks'].str.contains("Yard", na=False))]

        valve_data = filtered_df.groupby(["General Material Description"]).agg({
            "Quantity": "sum",
            "Weight": "sum"
        }).reset_index()

        valve_data_sbm = extract_df_sbm.groupby(["General Material Description"]).agg({
            "Quantity": "sum",
            "Weight": "sum"
        }).reset_index()

        valve_data_yard = extract_df_yard.groupby(["General Material Description"]).agg({
            "Quantity": "sum",
            "Weight": "sum"
        }).reset_index()

        return valve_data, valve_data_sbm, valve_data_yard
    else:
        raise ValueError(f"No data found for project number '{project_number}'.")



def get_bolt_mto_data(project_number):
    material_type = "Bolt"

    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if material_type in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    # Set the columns to extract
    columns_to_extract = ["Project Number", "Pipe Base Material", "SBM scope", "Total QTY to commit", "Quantity UOM"]

    # Check if the project number matches the specified one
    if df["Project Number"].astype(str).str.contains(str(project_number)).any():
        # Find the specified columns
        extract_columns = []
        for column in df.columns:
            if any(col in str(column) for col in columns_to_extract):
                extract_columns.append(column)

        # Filter the DataFrame and extract the desired data
        filtered_df = df[extract_columns]

        # Filter based on SBM scope and notnull values
        extract_df_sbm = filtered_df[(filtered_df['SBM scope'] == True) & (filtered_df['SBM scope'].notnull())]
        extract_df_yard = filtered_df[(filtered_df['SBM scope'] == False) & (filtered_df['SBM scope'].notnull())]

        bolt_data = filtered_df.groupby(["Pipe Base Material"]).agg({
            "Total QTY to commit": "sum",
            "Quantity UOM": "PCE"
        }).reset_index()

        bolt_sbm_data = extract_df_sbm.groupby(["Pipe Base Material"]).agg({
            "Total QTY to commit": "sum",
            "Quantity UOM": "PCE"
        }).reset_index()

        bolt_data_yard = extract_df_yard.groupby(["Pipe Base Material"]).agg({
            "Total QTY to commit": "sum",
            "Quantity UOM": "PCE"
        }).reset_index()

        return bolt_data, bolt_sbm_data, bolt_data_yard

    else:
        raise ValueError(f"No data found for project number '{project_number}'.")


def get_bent_mto_data():
    # TODO: Implement this function
    pass

def get_structure_mto_data():
    # TODO: Implement this function
    pass

def get_specialpip_mto_data():
    # TODO: Implement this function
    pass

def export_complete_mto_excel(data_frame):
    # TODO: Implement this function
    pass

def export_complete_mto_graphics(data_frame):
    # TODO: Implement this function
    pass

def export_complete_mto_pdf(data_frame):
    # TODO: Implement this function
    pass