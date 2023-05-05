import os
import pandas as pd
from datetime import datetime


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

#                       ------------------------ Piping --------------------


def extract_distinct_tag_numbers_piping(folder_path, project_number, material_type):
    excel_files = [
        f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    matching_files = [
        f for f in excel_files if str(project_number) in f and material_type in f
    ]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found."
        )

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    pipe_base_materials = set()
    tag_numbers = {}
    material_info = {}

    required_columns = ["Pipe Base Material", "Tag Number", "Total QTY to commit", "Unit Weight", "Total NET weight", "Quantity UOM", "Unit Weight UOM"]
    if all(col in df.columns for col in required_columns):
        materials = df["Pipe Base Material"].unique()

        for material in materials:
            pipe_base_materials.add(material)
            material_rows = df[df["Pipe Base Material"] == material]

            for _, row in material_rows.iterrows():
                tag_number = row["Tag Number"]

                # Store tag_number directly in the tag_numbers dictionary
                tag_numbers[tag_number] = material

                total_qty_to_commit = row["Total QTY to commit"]
                unit_weight = row["Unit Weight"]
                total_net_weight = row["Total NET weight"]
                quantity_uom = row["Quantity UOM"]
                unit_weight_uom = row["Unit Weight UOM"]

                # Store material information directly in the material_info dictionary based on tag_number
                material_info[tag_number] = {
                    "Total QTY to commit": total_qty_to_commit,
                    "Unit Weight": unit_weight,
                    "Total NET weight": total_net_weight,
                    "Quantity UOM": quantity_uom,
                    "Unit Weight UOM": unit_weight_uom
                }

    material_cost_analyze_piping_by_tag(project_number, tag_numbers, material_info)
    material_currency_cost_analyze_piping_by_tag(project_number, tag_numbers, material_info)

    return list(pipe_base_materials), tag_numbers, material_info


def material_cost_analyze_piping_by_tag(project_number, tag_numbers, material_info):
    folder_path = "../Data Pool/Ecosys API Data/PO Lines"

    excel_files = [
        f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    matching_files = [
        f for f in excel_files if str(project_number) in f
    ]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found."
        )

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    cost_data = []
    unmatched_data = []

    for tag, material in tag_numbers.items():
        tag_rows = df[df["Tag Number"] == tag]

        if tag_rows.empty:
            unmatched_data.append({
                "Project Number": project_number,
                "Base Material": material,
                "Tag Number": tag
            })
        else:
            tag_info = material_info[tag]
            total_cost = tag_rows["Cost Project Currency"].sum()

            cost_data.append({
                "Project Number": project_number,
                "Tag Number": tag,
                "Base Material": material,
                "Cost": total_cost,
                "Currency": "USD",
                "Total QTY to commit": tag_info["Total QTY to commit"],
                "Unit Weight": tag_info["Unit Weight"],
                "Total NET weight": tag_info["Total NET weight"],
                "Quantity UOM": tag_info["Quantity UOM"],
                "Unit Weight UOM": tag_info["Unit Weight UOM"]
            })

    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    result_folder_path = "../Data Pool/DCT Process Results"
    cost_df = pd.DataFrame(cost_data)
    output_file = os.path.join(result_folder_path,
                               f"MP{project_number}_Piping_TagBased_Cost_Analyze_{timestamp}.xlsx")
    cost_df.to_excel(output_file, index=False)

    # Save the unmatched data to a new Excel file
    if unmatched_data:
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        unmatched_df = pd.DataFrame(unmatched_data)
        output_file_unmatched = os.path.join(result_folder_path, f"Piping_NotMatched_TagNumber_{timestamp}.xlsx")
        unmatched_df.to_excel(output_file_unmatched, index=False)


def material_currency_cost_analyze_piping_by_tag(project_number, tag_numbers, material_info):
    folder_path = "../Data Pool/Ecosys API Data/PO Lines"

    excel_files = [
        f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    matching_files = [
        f for f in excel_files if str(project_number) in f
    ]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found."
        )

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    cost_data = []

    for tag, material in tag_numbers.items():
        tag_rows = df[df["Tag Number"] == tag]

        if tag_rows.empty:
            continue

        tag_info = material_info[tag]
        currency_groups = tag_rows.groupby("Transaction Currency")
        for currency, group in currency_groups:
            cost_data.append({
                "Project Number": project_number,
                "Tag Number": tag,
                "Base Material": material,
                "Cost": group["Cost Transaction Currency"].sum(),
                "Transaction Currency": currency,
                "Total QTY to commit": tag_info["Total QTY to commit"],
                "Unit Weight": tag_info["Unit Weight"],
                "Total NET weight": tag_info["Total NET weight"],
                "Quantity UOM": tag_info["Quantity UOM"],
                "Unit Weight UOM": tag_info["Unit Weight UOM"]
            })

    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    result_folder_path = "../Data Pool/DCT Process Results"
    cost_df = pd.DataFrame(cost_data)
    output_file = os.path.join(result_folder_path,
                               f"MP{project_number}_Piping_TagCurrency_Cost_Analyze_{timestamp}.xlsx")
    cost_df.to_excel(output_file, index=False)

