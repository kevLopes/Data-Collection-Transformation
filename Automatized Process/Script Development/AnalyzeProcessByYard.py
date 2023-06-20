import os
import pandas as pd
from datetime import datetime
import ExportReportsGraphics


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


def extract_distinct_product_codes_piping_yard(folder_path, project_number, material_type):
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

    required_columns = ["Pipe Base Material", "Product Code", "Total QTY to commit", "Unit Weight", "Total NET weight", "Quantity UOM", "Unit Weight UOM"]
    if all(col in df.columns for col in required_columns):
        materials = df["Pipe Base Material"].unique()

        for material in materials:
            pipe_base_materials.add(material)
            material_rows = df[df["Pipe Base Material"] == material]
            product_codes = set()

            for code in material_rows["Product Code"]:
                index = code.rfind(".")
                if index != -1:
                    product_codes.add(code[:index])

            material_codes[material] = product_codes

            total_qty_to_commit = material_rows["Total QTY to commit"].sum()
            qty_and_uom = material_rows.groupby("Quantity UOM")["Total QTY to commit"].sum().to_dict()
            unit_weight = material_rows["Unit Weight"].mean()
            total_net_weight = material_rows["Total NET weight"].sum()
            average_net_weight = material_rows["Total NET weight"].mean()
            unit_weight_uom = material_rows["Unit Weight UOM"].unique()

            material_info[material] = {
                "Total QTY commit per UOM": qty_and_uom,
                "Unit Weight": unit_weight,
                "Total NET weight": total_net_weight,
                "Average Net Weight": average_net_weight,
                "Unit Weight UOM": unit_weight_uom
            }

    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")