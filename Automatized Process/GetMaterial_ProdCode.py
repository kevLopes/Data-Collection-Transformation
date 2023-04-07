import os
import pandas as pd
from datetime import datetime
import time


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

    for file in matching_files:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)

        if "Pipe Base Material" not in df.columns or "Product Code" not in df.columns:
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

    folder_path = "C:\\Users\\keven.deoliveiralope\\Documents\Data Analyze Automatization\\Scripts\\Data-Collection-Transformation-kevLopes-DCT\\Automatized Process\\Data Pool\\Ecosys API Data\\PO Lines"
    material_cost_analyze(folder_path, project_number, material, material_codes)

    return list(pipe_base_materials), material_codes



def material_cost_analyze(folder_path, project_number, pipe_base_materials, material_codes):
    excel_files = [
        f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    cost_data = []

    for file in excel_files:
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
                cost_data.append({"Project Number": project_number, "Base Material": material, "Cost": material_cost, "Hours": material_hours})

    # Get the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    cost_df = pd.DataFrame(cost_data)
    output_file = os.path.join(folder_path, f"Material_Cost_Analyze_{timestamp}.xlsx")
    cost_df.to_excel(output_file, index=False)



