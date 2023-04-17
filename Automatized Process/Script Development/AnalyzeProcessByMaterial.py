import os
import pandas as pd
from datetime import datetime

def extract_distinct_product_codes(folder_path, project_number, material_type):
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

    pipe_base_materials = set()
    material_codes = {}
    material_info = {}

    for file in matching_files:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)

        required_columns = ["Pipe Base Material", "Product Code", "Total QTY to commit", "Unit Weight", "Total NET weight"]
        if not all(col in df.columns for col in required_columns):
            continue

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

            material_info[material] = {
                "Total QTY to commit": total_qty_to_commit,
                "Unit Weight": unit_weight,
                "Total NET weight": total_net_weight
            }

    material_cost_analyze(project_number, material, material_codes, material_info)

    return list(pipe_base_materials), material_codes, material_info


def material_cost_analyze(project_number, pipe_base_materials, material_codes, material_info):
    #folder Path
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

    cost_data = []

    for file in matching_files:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)

        if "Product Code" not in df.columns or "Cost" not in df.columns or "Hours" not in df.columns:
            continue

        for material, codes in material_codes.items():
            material_cost = 0
            material_hours = 0

            for code in codes:
                cost_rows = df[df["Product Code"].apply(lambda x: isinstance(x, str) and x.startswith(code))]
                material_cost += cost_rows["Cost"].sum()
                material_hours += cost_rows["Hours"].sum()

            if material_cost > 0 or material_hours > 0:
                cost_data.append({
                    "Project Number": project_number,
                    "Base Material": material,
                    "Cost": material_cost,
                    "Hours": material_hours,
                    "Total QTY to commit": material_info[material]["Total QTY to commit"],
                    "Unit Weight": material_info[material]["Unit Weight"],
                    "Total NET weight": material_info[material]["Total NET weight"]
                })

    # Get the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    result_folder_path = "../Data Pool/DCT Process Results"
    cost_df = pd.DataFrame(cost_data)
    output_file = os.path.join(result_folder_path, f"{project_number}_Material_Cost_Analyze_{timestamp}.xlsx")
    cost_df.to_excel(output_file, index=False)
