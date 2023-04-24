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


def extract_distinct_product_codes_piping(folder_path, project_number, material_type):
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
    material_codes = {}
    material_info = {}

    required_columns = ["Pipe Base Material", "Product Code", "Total QTY to commit", "Unit Weight", "Total NET weight"]
    if all(col in df.columns for col in required_columns):
        materials = df["Pipe Base Material"].unique()

        for material in materials:
            pipe_base_materials.add(material)
            material_rows = df[df["Pipe Base Material"] == material]
            product_codes = set()  # Create a new set for each material

            for code in material_rows["Product Code"]:
                index = code.rfind(".")
                if index != -1:
                    product_codes.add(code[:index])

            material_codes[material] = product_codes

            # Extract additional columns data for each material
            total_qty_to_commit = material_rows["Total QTY to commit"].sum()
            unit_weight = material_rows["Unit Weight"].mean()
            total_net_weight = material_rows["Total NET weight"].sum()
            average_net_weight = material_rows["Total NET weight"].mean()

            material_info[material] = {
                "Total QTY to commit": total_qty_to_commit,
                "Unit Weight": unit_weight,
                "Total NET weight": total_net_weight,
                "Average Net Weight": average_net_weight
            }

    material_cost_analyze_piping(project_number, material_codes, material_info)

    return list(pipe_base_materials), material_codes, material_info


def material_cost_analyze_piping(project_number, material_codes, material_info):
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

    if "Product Code" in df.columns and "Cost Project Currency" in df.columns and "Quantity" in df.columns:
        for material, codes in material_codes.items():
            material_cost = 0
            material_quantity = 0

            for code in codes:
                cost_rows = df[df["Product Code"].apply(lambda x: isinstance(x, str) and x.startswith(code))]
                material_cost += cost_rows["Cost Project Currency"].sum()
                material_quantity += cost_rows["Quantity"].sum()

            if material_cost > 0 or material_quantity > 0:
                cost_data.append({
                    "Project Number": project_number,
                    "Base Material": material,
                    "Product Code": ", ".join(codes),
                    "Cost": material_cost,
                    "Quantity": material_quantity,
                    "Total QTY to commit": material_info[material]["Total QTY to commit"],
                    "Unit Weight": material_info[material]["Unit Weight"],
                    "Total NET weight": material_info[material]["Total NET weight"],
                    "Average Net Weight": material_info[material]["Average Net Weight"]
                })

    # Get the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    result_folder_path = "../Data Pool/DCT Process Results"
    cost_df = pd.DataFrame(cost_data)
    output_file = os.path.join(result_folder_path, f"MP{project_number}_Piping_MaterialBased_Cost_Analyze_{timestamp}.xlsx")
    cost_df.to_excel(output_file, index=False)

#                       ------------------------ Valve --------------------


def extract_distinct_product_codes_valve(folder_path, project_number, material_type):
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

    general_materials = set()
    material_codes = {}
    material_info = {}

    required_columns = ["General Material Description", "Product Code", "Quantity", "Weight", "SIZE (inch)"]
    if all(col in df.columns for col in required_columns):
        materials = df["General Material Description"].unique()

        for material in materials:
            general_materials.add(material)
            material_rows = df[df["General Material Description"] == material]
            #product_codes = set()  # Create a new set for each material

            product_codes = set(material_rows["Product Code"].unique())

            material_codes[material] = product_codes

            # Extract additional columns data for each material
            total_qty = material_rows["Quantity"].sum()
            total_size = material_rows["SIZE (inch)"].sum()
            average_size = material_rows["SIZE (inch)"].mean()
            total_weight = material_rows["Weight"].sum()
            average_weight = material_rows["Weight"].mean()

            material_info[material] = {
                "Quantity": total_qty,
                "SIZE (inch)": total_size,
                "Weight": total_weight,
                "Average Weight": average_weight,
                "Average Size (inch)": average_size
            }

    material_cost_analyze_valve(project_number, material_codes, material_info)

    return list(general_materials), material_codes, material_info


def material_cost_analyze_valve(project_number, material_codes, material_info):
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

    if "Product Code" in df.columns and "Cost Project Currency" in df.columns:
        for material, codes in material_codes.items():
            material_cost = 0

            for code in codes:
                cost_rows = df[df["Product Code"].apply(lambda x: isinstance(x, str) and x.startswith(code))]
                material_cost += cost_rows["Cost Project Currency"].sum()

            if material_cost > 0:
                cost_data.append({
                    "Project Number": project_number,
                    "Base Material": material,
                    "Product Code": ", ".join(codes),
                    "Cost": material_cost,
                    "Quantity": material_info[material]["Quantity"],
                    "SIZE (inch)": material_info[material]["SIZE (inch)"],
                    "Average Size (inch)": material_info[material]["Average Size (inch)"],
                    "Weight": material_info[material]["Weight"],
                    "Average Weight": material_info[material]["Average Weight"]
                })

    # Get the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    result_folder_path = "../Data Pool/DCT Process Results"
    cost_df = pd.DataFrame(cost_data)
    output_file = os.path.join(result_folder_path, f"MP{project_number}_Valve_MaterialBased_Cost_Analyze_{timestamp}.xlsx")
    cost_df.to_excel(output_file, index=False)

#                       ------------------------ Bolt --------------------


def extract_distinct_product_codes_bolt(folder_path, project_number, material_type):
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

    general_materials = set()
    material_codes = {}
    material_info = {}

    required_columns = ["Pipe Base Material", "Product Code", "Total QTY to commit", "Qty confirmed in design", "SIZE", "SBM scope"]
    if all(col in df.columns for col in required_columns):
        materials = df["Pipe Base Material"].unique()

        for material in materials:
            general_materials.add(material)
            material_rows = df[df["Pipe Base Material"] == material]
            product_codes = set()  # Create a new set for each material

            for code in material_rows["Product Code"]:
                index = code.rfind(".")
                if index != -1:
                    product_codes.add(code[:index])

            material_codes[material] = product_codes

            # Extract additional columns data for each material
            total_qty_to_commit = material_rows["Total QTY to commit"].sum()
            #total_size = material_rows["SIZE"].sum()
            #average_size = material_rows["SIZE"].mean()
            total_qty_design = material_rows["Qty confirmed in design"].sum()
            #uom_weight = material_rows["Unit Weight UOM"].mean()

            material_info[material] = {
                "Total QTY to commit": total_qty_to_commit,
                #"SIZE": total_size,
                "Qty confirmed in design": total_qty_design
                #"Unit Weight UOM": uom_weight
                #"Average Size": average_size
            }

    material_cost_analyze_bolt(project_number, material_codes, material_info)

    return list(general_materials), material_codes, material_info


def material_cost_analyze_bolt(project_number, material_codes, material_info):
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

    if "Product Code" in df.columns and "Cost Project Currency" in df.columns and "Quantity" in df.columns:
        for material, codes in material_codes.items():
            material_cost = 0
            material_quantity = 0

            for code in codes:
                cost_rows = df[df["Product Code"].apply(lambda x: isinstance(x, str) and x.startswith(code))]
                material_cost += cost_rows["Cost Project Currency"].sum()
                material_quantity += cost_rows["Quantity"].sum()

            if material_cost > 0 or material_quantity > 0:
                cost_data.append({
                    "Project Number": project_number,
                    "Base Material": material,
                    "Product Code": ", ".join(codes),
                    "Cost": material_cost,
                    "Quantity from POs": material_quantity,
                    "Total QTY to commit": material_info[material]["Total QTY to commit"],
                    #"SIZE": material_info[material]["SIZE"],
                    "Qty confirmed in design": material_info[material]["Qty confirmed in design"],
                    #"Unit Weight UOM": material_info[material]["Unit Weight UOM"
                    #"Average Size": material_info[material]["Average Size"]
                })

    # Get the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    result_folder_path = "../Data Pool/DCT Process Results"
    cost_df = pd.DataFrame(cost_data)
    output_file = os.path.join(result_folder_path, f"MP{project_number}_Bolt_MaterialBased_Cost_Analyze_{timestamp}.xlsx")
    cost_df.to_excel(output_file, index=False)

