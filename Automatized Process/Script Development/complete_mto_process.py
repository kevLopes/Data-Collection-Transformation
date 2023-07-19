import os
import numpy as np
import pandas as pd
from datetime import datetime

folder_path = "../Data Pool/Data Hub Materials"


#Complete analyze of those different MTO files - Material type, Weight, Quantities
def complete_mto_data_analyze(project_number):

    #Call each function and collect their data
    piping_data = get_piping_mto_data(project_number)
    valve_data = get_valve_mto_data(project_number)
    bolt_data = get_bolt_mto_data(project_number)
    structure_data = get_structure_mto_data(project_number)
    specialpip_data = get_specialpip_mto_data(project_number)

    #Piping Extra Data details
    piping_extra_detail_data = get_piping_extra_details(project_number, "Piping")
    #Valve Extra Data details
    valve_extra_detail_data = get_valve_extra_details(project_number, "Valve")
    #Bolt Extra Data details
    bolt_extra_detail_data = get_bolt_extra_details(project_number, "Bolt")
    #Special Piping Extra details
    spc_piping_extra_detail_data = get_spc_piping_extra_details(project_number, "Piping")

    # Combine the data into a data frame
    data_frame = pd.DataFrame({
        "Piping Data": piping_data,
        "Valve Data": valve_data,
        "Bolt Data": bolt_data,
        "Structure Data": structure_data,
        "Special Piping Data": specialpip_data
    })

    # Pass the data frame to export functions
    export_complete_mto_excel()
    export_complete_mto_graphics()
    export_complete_mto_pdf()


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


#PIPING MTO DATA
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


#VALVE MTO DATA
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


#BOLT MTO DATA
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
            "Total QTY to commit": "sum"
        }).reset_index()

        bolt_sbm_data = extract_df_sbm.groupby(["Pipe Base Material"]).agg({
            "Total QTY to commit": "sum"
        }).reset_index()

        bolt_data_yard = extract_df_yard.groupby(["Pipe Base Material"]).agg({
            "Total QTY to commit": "sum"
        }).reset_index()

        return bolt_data, bolt_sbm_data, bolt_data_yard

    else:
        raise ValueError(f"No data found for project number '{project_number}'.")


#STRUCTURE MTO DATA
def get_structure_mto_data(project_number):
    material_type = "Structure"
    folder_path_strct = "../Data Pool/DCT Process Results/Exported Result Files/Structure"

    excel_files = [f for f in os.listdir(folder_path_strct) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if str(project_number) in f and material_type in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

    most_recent_file = get_most_recent_file(folder_path_strct, matching_files)
    file_path = os.path.join(folder_path_strct, most_recent_file)
    df = pd.read_excel(file_path)

    structure_qty_pce = df[(df['Quantity UOM'] == "PCS")]
    structure_qty_m = df[(df['Quantity UOM'] == "m")]
    structure_qty_m2 = df[(df['Quantity UOM'] == "m2")]

    structure_totals_pcs = structure_qty_pce.groupby(["Quantity UOM"]).agg({
        'Total QTY to commit': 'sum',
        'Total NET weight': 'sum',
        'Unit Weight': 'mean',
        'Thickness': 'sum',
        'Wastage Quantity': 'sum',
        'Quantity Including Wastage': 'sum',
        'Total Gross Weight': 'sum', }).reset_index()

    structure_totals_m = structure_qty_m.groupby(["Quantity UOM"]).agg({
        'Total QTY to commit': 'sum',
        'Total NET weight': 'sum',
        'Unit Weight': 'mean',
        'Thickness': 'sum',
        'Wastage Quantity': 'sum',
        'Quantity Including Wastage': 'sum',
        'Total Gross Weight': 'sum', }).reset_index()

    structure_totals_m2 = structure_qty_m2.groupby(["Quantity UOM"]).agg({
        'Total QTY to commit': 'sum',
        'Total NET weight': 'sum',
        'Unit Weight': 'mean',
        'Thickness': 'sum',
        'Wastage Quantity': 'sum',
        'Quantity Including Wastage': 'sum',
        'Total Gross Weight': 'sum', }).reset_index()

    return structure_totals_m2, structure_totals_m, structure_totals_pcs


#SPECIAL PIPING MTO DATA
def get_specialpip_mto_data(project_number):
    material_type = "Special PIP"

    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if material_type in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    # Set the columns to extract
    columns_to_extract = ["TagNumber", "Size (Inch)", "Weight", "Remarks", "PO Number", "Qty"]

    # Find the specified columns
    extract_columns = []
    for column in df.columns:
        if any(col in str(column) for col in columns_to_extract):
            extract_columns.append(column)

    # Filter the DataFrame and extract the desired data
    filtered_df = df.loc[:, extract_columns].copy()  # Use .loc with [:] to select all rows and create a copy

    # Convert Qty and Weight columns to numeric type
    filtered_df.loc[:, "Qty"] = pd.to_numeric(filtered_df["Qty"], errors="coerce")
    filtered_df.loc[:, "Weight"] = pd.to_numeric(filtered_df["Weight"], errors="coerce")

    # Filter based on SBM scope and notnull values
    extract_df_sbm = filtered_df[(filtered_df['PO Number'] != "BY YARD") & (filtered_df['PO Number'].notnull())]
    extract_df_yard = filtered_df[(filtered_df['PO Number'] == "BY YARD") & (filtered_df['PO Number'].notnull())]

    spcpip_data = filtered_df.agg({
        "Qty": "sum",
        "Weight": "sum"
    }).reset_index()

    spcpip_sbm_data = extract_df_sbm.agg({
        "Qty": "sum",
        "Weight": "sum"
    }).reset_index()

    spcpip_data_yard = extract_df_yard.agg({
        "Qty": "sum",
        "Weight": "sum"
    }).reset_index()

    return spcpip_data, spcpip_data_yard, spcpip_sbm_data


#function to read more information from Piping
def get_piping_extra_details(project_number, material_type):
    first_folder_path = "../Data Pool/Material Data Organized/Piping"
    excel_files = [
        f for f in os.listdir(first_folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    matching_files = [
        f for f in excel_files if str(project_number) in f and material_type in f
    ]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found."
        )

    most_recent_file = get_most_recent_file(first_folder_path, matching_files)
    file_path = os.path.join(first_folder_path, most_recent_file)
    first_df = pd.read_excel(file_path)

    # Collect tag numbers from the first file
    first_tag_numbers = first_df["Tag Number"].unique()

    # Create an empty dictionary to store tag numbers and their details
    tag_numbers = {}

    # Iterate over first_tag_numbers and store tag numbers and their details
    for tag_number in first_tag_numbers:
        tag_numbers[tag_number] = {
            "Pipe Base Material": first_df[first_df["Tag Number"] == tag_number]["Pipe Base Material"].iloc[0],
            "Unit Weight": first_df[first_df["Tag Number"] == tag_number]["Unit Weight"].iloc[0]
        }

    # Search for the tag numbers in the second file
    second_file_folder = "../Data Pool/Ecosys API Data/PO Lines"
    second_excel_files = [
        f for f in os.listdir(second_file_folder) if f.endswith(".xlsx") or f.endswith(".xls")
    ]
    matching_second_files = [f for f in second_excel_files if str(project_number) in f]

    if not matching_second_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found in the PO Lines folder."
        )

    second_most_recent_file = get_most_recent_file(second_file_folder, matching_second_files)
    second_file_path = os.path.join(second_file_folder, second_most_recent_file)
    second_df = pd.read_excel(second_file_path)

    cost_data = []

    for tag_number, material_info in tag_numbers.items():
        regular_rows = second_df[second_df["Tag Number"] == tag_number]

        if not regular_rows.empty:
            surplus_rows = second_df[
                second_df["Tag Number"].str.startswith(tag_number) &
                (second_df["Tag Number"].str.endswith("-SURPLUS") | second_df["Tag Number"].str.endswith("-SURPLUS+"))
                ]

            combined_rows = pd.concat([regular_rows, surplus_rows])

            if combined_rows.empty:
                continue

            # Perform calculations and logic based on combined_rows
            calc_weight = 0.0
            po_quantity_by_uom = {}
            total_cost = 0.0

            for _, row in combined_rows.iterrows():
                uom = row["UOM"]
                quantity = row["Quantity"]

                # Group the quantities by UOM
                if uom not in po_quantity_by_uom:
                    po_quantity_by_uom[uom] = 0.0
                po_quantity_by_uom[uom] += quantity
                calc_weight += quantity * material_info["Unit Weight"]

                total_cost += row["Cost Project Currency"]

            # Store the calculated values
            cost_data.append({
                "Tag Number": tag_number,
                "Pipe Base Material": material_info["Pipe Base Material"],
                "PO Quantity by UOM": po_quantity_by_uom,
                "PO Weight": calc_weight,
                "Total Cost": total_cost
            })

    # Calculate totals
    total_matched_tags = len(cost_data)
    total_unmatched_tags = len(tag_numbers) - len(cost_data)
    total_surplus_tags = second_df["Tag Number"].str.contains("-SURPLUS|SURPLUS+S", na=False).sum()
    total_weight = sum(entry["PO Weight"] for entry in cost_data)
    total_quantity_by_uom = {}

    for entry in cost_data:
        for uom, quantity in entry["PO Quantity by UOM"].items():
            if uom not in total_quantity_by_uom:
                total_quantity_by_uom[uom] = 0.0
            total_quantity_by_uom[uom] += quantity

    total_cost = sum(entry["Total Cost"] for entry in cost_data)

    # Calculate total cost by material
    total_cost_by_material = {}
    overall_cost = total_cost

    for entry in cost_data:
        material = entry["Pipe Base Material"]
        total_cost = entry["Total Cost"]

        if material not in total_cost_by_material:
            total_cost_by_material[material] = 0.0
        total_cost_by_material[material] += total_cost

    # Extract unique Cost Object IDs from the second file for matched tag numbers
    matched_tag_numbers = [entry["Tag Number"] for entry in cost_data]
    unique_cost_object_ids = second_df.loc[
        second_df["Tag Number"].isin(matched_tag_numbers), "Cost Object ID"
    ].unique().tolist()

    # Calculate the total cost for SURPLUS tag numbers
    total_surplus_cost = second_df.loc[
        second_df["Tag Number"].str.endswith("-SURPLUS") | second_df["Tag Number"].str.endswith("-SURPLUS+"),
        "Cost Project Currency"
    ].sum()

    # Extract unique Cost Object IDs from the second file for SURPLUS tag numbers
    surplus_tag_numbers = second_df.loc[
        second_df["Tag Number"].str.endswith("-SURPLUS") | second_df["Tag Number"].str.endswith("-SURPLUS+"),
        "Tag Number"
    ].unique().tolist()
    unique_surplus_cost_object_ids = second_df.loc[
        second_df["Tag Number"].isin(surplus_tag_numbers), "Cost Object ID"
    ].unique().tolist()

    # Return the calculated data along with the unique Cost Object IDs, total surplus cost, and unique surplus Cost Object IDs
    return (
        total_matched_tags,
        total_unmatched_tags,
        total_surplus_tags,
        total_weight,
        total_quantity_by_uom,
        overall_cost,
        total_cost_by_material,
        unique_cost_object_ids,
        total_surplus_cost,
        unique_surplus_cost_object_ids
    )


#function to read more information from Special Piping
def get_spc_piping_extra_details(project_number, material_type):
    first_folder_path = "../Data Pool/Material Data Organized/Special Piping"
    excel_files = [f for f in os.listdir(first_folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if str(project_number) in f and material_type in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found."
        )

    most_recent_file = get_most_recent_file(first_folder_path, matching_files)
    file_path = os.path.join(first_folder_path, most_recent_file)
    first_df = pd.read_excel(file_path)

    # Collect tag numbers from the first file
    first_tag_numbers = first_df["TagNumber"].unique()

    # Search for the tag numbers in the second file
    second_file_folder = "../Data Pool/Ecosys API Data/PO Lines"
    second_excel_files = [f for f in os.listdir(second_file_folder) if f.endswith(".xlsx") or f.endswith(".xls")]
    matching_second_files = [f for f in second_excel_files if str(project_number) in f]

    if not matching_second_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found in the PO Lines folder.")

    second_most_recent_file = get_most_recent_file(second_file_folder, matching_second_files)
    second_file_path = os.path.join(second_file_folder, second_most_recent_file)
    second_df = pd.read_excel(second_file_path)

    # Iterate over the tag numbers
    cost_data = []
    for tag_number in np.nditer(first_tag_numbers, flags=['refs_ok']):
        regular_rows = second_df[second_df["Tag Number"] == tag_number]

        if not regular_rows.empty:
            # Perform calculations and logic based on regular_rows
            quantity = 0.0
            total_cost = 0.0

            for _, row in regular_rows.iterrows():
                quantity += row["Quantity"]
                total_cost += row["Cost Project Currency"]
                po = row["Cost Object ID"]

            # Store the calculated values
            cost_data.append({
                "Tag Number": tag_number,
                "PO Quantity": quantity,
                "Total Cost": total_cost,
                "PO Number": po
            })

    # Calculate totals
    total_matched_tags = len(cost_data)
    total_unmatched_tags = len(first_tag_numbers) - len(cost_data)
    total_quantity_by_uom = sum(entry["PO Quantity"] for entry in cost_data)
    total_cost = sum(entry["Total Cost"] for entry in cost_data)

    # Get distinct PO Numbers
    po_list = list(set(entry["PO Number"] for entry in cost_data))

    # Return the calculated data along with the unique Cost Object IDs, total surplus cost, and unique surplus Cost Object IDs
    return (
        total_matched_tags,
        total_unmatched_tags,
        total_quantity_by_uom,
        total_cost,
        po_list
    )


#function to read more information from Valve
def get_valve_extra_details(project_number, material_type):
    first_folder_path = "../Data Pool/Material Data Organized/Valve"
    excel_files = [
        f for f in os.listdir(first_folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    matching_files = [
        f for f in excel_files if str(project_number) in f and material_type in f
    ]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found."
        )

    most_recent_file = get_most_recent_file(first_folder_path, matching_files)
    file_path = os.path.join(first_folder_path, most_recent_file)
    first_df = pd.read_excel(file_path)

    tag_numbers = {}

    for _, row in first_df.iterrows():
        tag_number = row["Product Code"]
        general_description = row["General Material Description"]
        weight = row["Weight"]

        tag_numbers[tag_number] = {
            "General Material Description": general_description,
            "Weight": weight
        }

    second_file_folder = "../Data Pool/Ecosys API Data/PO Lines"
    second_excel_files = [
        f for f in os.listdir(second_file_folder) if f.endswith(".xlsx") or f.endswith(".xls")
    ]
    matching_second_files = [f for f in second_excel_files if str(project_number) in f]

    if not matching_second_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found in the PO Lines folder.")

    second_most_recent_file = get_most_recent_file(second_file_folder, matching_second_files)
    second_file_path = os.path.join(second_file_folder, second_most_recent_file)
    second_df = pd.read_excel(second_file_path)

    cost_data = []

    for _, row in second_df.iterrows():
        tag_number = row["Product Code"]

        if tag_number in tag_numbers:
            uom = row["UOM"]
            quantity = row["Quantity"]
            weight = tag_numbers[tag_number]["Weight"]

            calc_weight = quantity * weight

            total_cost = row["Cost Project Currency"]

            cost_data.append({
                "Product Code": tag_number,
                "General Material Description": tag_numbers[tag_number]["General Material Description"],
                "PO Quantity": quantity,
                "PO Weight": calc_weight,
                "Total Cost": total_cost
            })

    total_weight = sum(entry["PO Weight"] for entry in cost_data)
    total_cost = sum(entry["Total Cost"] for entry in cost_data)
    total_quantity = sum(entry["PO Quantity"] for entry in cost_data)

    overall_cost = total_cost

    # Calculate cost per General Description type
    cost_by_general_description = {}

    for entry in cost_data:
        general_description = entry["General Material Description"]
        total_cost = entry["Total Cost"]

        if general_description not in cost_by_general_description:
            cost_by_general_description[general_description] = 0.0
        cost_by_general_description[general_description] += total_cost

    return (
        total_quantity,
        overall_cost,
        cost_by_general_description
    )


#function to read more information from Bolt
def get_bolt_extra_details(project_number, material_type):
    first_folder_path = "../Data Pool/Material Data Organized/Bolt"
    excel_files = [f for f in os.listdir(first_folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if str(project_number) in f and material_type in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found."
        )

    most_recent_file = get_most_recent_file(first_folder_path, matching_files)
    file_path = os.path.join(first_folder_path, most_recent_file)
    first_df = pd.read_excel(file_path)

    # Collect product codes from the first file
    first_product_codes = first_df["Product Code"].unique()

    product_codes = {}

    # Iterate over first_product_codes and store product codes and their details
    for code in first_product_codes:
        product_codes[code] = {
            "Pipe Base Material": first_df[first_df["Product Code"] == code]["Pipe Base Material"].iloc[0],
            "Total QTY to commit": first_df[first_df["Product Code"] == code]["Total QTY to commit"].iloc[0]
        }

    second_file_folder = "../Data Pool/Ecosys API Data/PO Lines"
    second_excel_files = [f for f in os.listdir(second_file_folder) if f.endswith(".xlsx") or f.endswith(".xls")]
    matching_second_files = [f for f in second_excel_files if str(project_number) in f]

    if not matching_second_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found in the PO Lines folder."
        )

    second_most_recent_file = get_most_recent_file(second_file_folder, matching_second_files)
    second_file_path = os.path.join(second_file_folder, second_most_recent_file)
    second_df = pd.read_excel(second_file_path)

    cost_data = []

    for _, row in second_df.iterrows():
        code = row["Product Code"]

        if code in product_codes:
            quantity = row["Quantity"]
            total_cost = row["Cost Project Currency"]

            cost_data.append({
                "Product Code": code,
                "Pipe Base Material": product_codes[code]["Pipe Base Material"],
                "PO Quantity": quantity,
                "Total Cost": total_cost
            })

    total_cost = sum(entry["Total Cost"] for entry in cost_data)
    total_po_quantity = sum(entry["PO Quantity"] for entry in cost_data)

    overall_cost = total_cost

    # Calculate cost per Pipe Base Material
    cost_by_pipe_base_material = {}

    for entry in cost_data:
        pipe_base_material = entry["Pipe Base Material"]
        total_cost = entry["Total Cost"]

        if pipe_base_material not in cost_by_pipe_base_material:
            cost_by_pipe_base_material[pipe_base_material] = 0.0
        cost_by_pipe_base_material[pipe_base_material] += total_cost

    # Get the Product Codes from the first file that were not found in the second file
    missing_product_codes = set(product_codes.keys()) - set(entry["Product Code"] for entry in cost_data)

    return (
        total_po_quantity,
        overall_cost,
        cost_by_pipe_base_material,
        missing_product_codes
    )



def export_complete_mto_excel():
    # TODO: Implement this function
    pass


def export_complete_mto_graphics():
    # TODO: Implement this function
    pass


def export_complete_mto_pdf():
    # TODO: Implement this function
    pass
