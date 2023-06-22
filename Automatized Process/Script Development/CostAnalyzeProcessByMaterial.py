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

        material_cost_analyze_piping(project_number, material_codes, material_info)
        material_currency_cost_analyze_piping(project_number, material_codes, material_info)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")


def material_cost_analyze_piping(project_number, material_codes, material_info):
    folder_path = "../Data Pool/Ecosys API Data/PO Lines"

    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if str(project_number) in f]

    if not matching_files:
        raise FileNotFoundError(f"No files containing the project number '{project_number}' were found.")

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    cost_data = []
    unmatched_data = []

    if "Product Code" in df.columns and "Cost Project Currency" in df.columns and "Quantity" in df.columns and "UOM" in df.columns:
        for material, codes in material_codes.items():

            for code in codes:
                cost_rows = df[df["Product Code"].apply(lambda x: isinstance(x, str) and x.startswith(code))]

                # Check if the code has no match
                if cost_rows.empty:
                    unmatched_data.append({
                        "Project Number": project_number,
                        "Base Material": material,
                        "Product Code": code
                    })
                else:
                    for uom in cost_rows["UOM"].unique():
                        uom_rows = cost_rows[cost_rows["UOM"] == uom]
                        material_cost = uom_rows["Cost Project Currency"].sum()
                        material_quantity = uom_rows["Quantity"].sum()

                        if material_cost > 0 or material_quantity > 0:
                            cost_data.append({
                                "Project Number": project_number,
                                "Base Material": material,
                                "Product Code": code,
                                "Cost": material_cost,
                                "Currency": "USD",
                                "Total QTY to commit per UOM": material_info[material]["Total QTY commit per UOM"],
                                "Total NET weight": material_info[material]["Total NET weight"],
                                "Average Net Weight": material_info[material]["Average Net Weight"],
                                "Unit Weight UOM": material_info[material]["Unit Weight UOM"],
                                "Quantity in PO": material_quantity,
                                "UOM in PO": uom
                            })

        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Piping"
        cost_df = pd.DataFrame(cost_data)
        output_file = os.path.join(result_folder_path, f"MP{project_number}_Piping_MaterialBased_Cost_Analyze_{timestamp}.xlsx")
        cost_df.to_excel(output_file, index=False)

        # Save the unmatched data to a new Excel file
        if unmatched_data:
            timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
            unmatched_df = pd.DataFrame(unmatched_data)
            output_file_unmatched = os.path.join(result_folder_path, f"Piping_NotMatched_ProductCode_{timestamp}.xlsx")
            unmatched_df.to_excel(output_file_unmatched, index=False)

        cost_df_mt = cost_df.copy()
        cost_df_mw = cost_df.copy()

        ExportReportsGraphics.plot_piping_material_cost(cost_df_mt, project_number)
        ExportReportsGraphics.plot_piping_material_weight(cost_df_mw, project_number)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")



def material_currency_cost_analyze_piping(project_number, material_codes, material_info):
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

    required_columns = ["Product Code", "Quantity", "Cost Transaction Currency", "Transaction Currency", "UOM"]
    if all(column in df.columns for column in required_columns):
        for material, codes in material_codes.items():

            for code in codes:
                cost_rows = df[df["Product Code"].apply(lambda x: isinstance(x, str) and x.startswith(code))]

                if cost_rows.empty:
                    continue
                else:
                    for uom in cost_rows["UOM"].unique():
                        uom_rows = cost_rows[cost_rows["UOM"] == uom]
                        currency_groups = uom_rows.groupby("Transaction Currency")
                        for currency, group in currency_groups:
                            key = (currency, uom)
                            material_cost_by_currency = material_info[material].setdefault("Cost by Currency and UOM", {})
                            material_cost_by_currency[key] = material_cost_by_currency.get(key, 0) + group["Cost Transaction Currency"].sum()
                            material_quantity_by_currency = material_info[material].setdefault("Quantity by Currency and UOM", {})
                            material_quantity_by_currency[key] = material_quantity_by_currency.get(key, 0) + group["Quantity"].sum()

                    for (currency, uom), cost in material_info[material]["Cost by Currency and UOM"].items():
                        cost_data.append({
                            "Project Number": project_number,
                            "Base Material": material,
                            "Product Code": ", ".join(codes),
                            "Cost": cost,
                            "Transaction Currency": currency,
                            "Total QTY to commit per UOM": material_info[material]["Total QTY commit per UOM"],
                            "Total NET weight": material_info[material]["Total NET weight"],
                            "Average Net Weight": material_info[material]["Average Net Weight"],
                            "Weight UOM": material_info[material]["Unit Weight UOM"],
                            "Quantity in PO": material_info[material]["Quantity by Currency and UOM"][(currency, uom)],
                            "UOM in PO": uom
                        })

        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Piping"
        cost_df_c = pd.DataFrame(cost_data)
        output_file = os.path.join(result_folder_path,
                                   f"MP{project_number}_Piping_MatCurrency_CostAnalyze_{timestamp}.xlsx")
        cost_df_c.to_excel(output_file, index=False)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")

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

    required_columns = ["General Material Description", "Product Code", "Quantity", "Weight", "SIZE (inch)", "Status", "Remarks"]
    if all(col in df.columns for col in required_columns):
        filtered_df = df[(df["Status"] == "CONFIRMED") & (df["Remarks"] != "Yard Scope")]
        materials = filtered_df["General Material Description"].unique()

        general_materials = set()
        material_codes = {}
        material_info = {}

        for material in materials:
            general_materials.add(material)
            material_rows = filtered_df[filtered_df["General Material Description"] == material]

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
        material_currency_cost_analyze_valve(project_number, material_codes, material_info)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")



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

    cost_data_dict = {}
    unmatched_data = []

    if "Product Code" in df.columns and "Cost Project Currency" in df.columns and "Quantity" in df.columns and "UOM" in df.columns:
        for material, codes in material_codes.items():
            material_cost = 0
            quantity_po = 0

            for code in codes:
                cost_rows = df[df["Product Code"].apply(lambda x: isinstance(x, str) and x.startswith(code))]

                # Check if the code has no match
                if cost_rows.empty:
                    unmatched_data.append({
                        "Project Number": project_number,
                        "Base Material": material,
                        "Product Code": code
                    })
                else:
                    material_cost += cost_rows["Cost Project Currency"].sum()
                    quantity_po += cost_rows["Quantity"].sum()

            if material in cost_data_dict:
                cost_data_dict[material]["Cost"] += material_cost
                cost_data_dict[material]["Product Code"].extend(codes)
                cost_data_dict[material]["Quantity"] += material_info[material]["Quantity"]
                cost_data_dict[material]["Weight"] += material_info[material]["Weight"]
                cost_data_dict[material]["SIZE (inch)"] += material_info[material]["SIZE (inch)"]
                cost_data_dict[material]["Quantity in PO"] += quantity_po
                cost_data_dict[material]["Count"] += 1  # keep track of items count
            else:
                if material_cost > 0:
                    cost_data_dict[material] = {
                        "Project Number": project_number,
                        "Base Material": material,
                        "Product Code": codes,
                        "Cost": material_cost,
                        "Currency": "USD",
                        "Quantity": material_info[material]["Quantity"],
                        "SIZE (inch)": material_info[material]["SIZE (inch)"],
                        "Weight": material_info[material]["Weight"],
                        "Quantity in PO": quantity_po,
                        "Count": 1  # initialize count
                    }

        # Calculating averages and converting product codes to strings
        for material, data in cost_data_dict.items():
            data["Average Size (inch)"] = data["SIZE (inch)"] / data["Count"]
            data["Average Weight"] = data["Weight"] / data["Count"]
            data["Product Code"] = ", ".join(data["Product Code"])
            del data["Count"]  # remove count before exporting data

        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Valve"
        cost_data = list(cost_data_dict.values())
        cost_df = pd.DataFrame(cost_data)
        output_file = os.path.join(result_folder_path, f"MP{project_number}_Valve_MaterialBased_Cost_Analyze_{timestamp}.xlsx")
        cost_df.to_excel(output_file, index=False)

        # Save the unmatched data to a new Excel file
        if unmatched_data:
            timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
            unmatched_df = pd.DataFrame(unmatched_data)
            output_file_unmatched = os.path.join(result_folder_path, f"Valve_NotMatched_ProductCode_{timestamp}.xlsx")
            unmatched_df.to_excel(output_file_unmatched, index=False)

        ExportReportsGraphics.plot_valve_material_cost(cost_df, project_number)
        ExportReportsGraphics.plot_valve_material_quantity_weight(cost_df, project_number)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")


def material_currency_cost_analyze_valve(project_number, material_codes, material_info):
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

    required_columns = ["Product Code", "Quantity", "Cost Transaction Currency", "Transaction Currency"]
    if all(column in df.columns for column in required_columns):
        for material, codes in material_codes.items():
            material_cost_by_currency = {}
            material_quantity_by_currency = {}

            for code in codes:
                cost_rows = df[df["Product Code"].apply(lambda x: isinstance(x, str) and x.startswith(code))]

                if cost_rows.empty:
                    unmatched_data.append({
                        "Project Number": project_number,
                        "Base Material": material,
                        "Product Code": code
                    })
                else:
                    currency_groups = cost_rows.groupby("Transaction Currency")
                    for currency, group in currency_groups:
                        material_cost_by_currency[currency] = material_cost_by_currency.get(currency, 0) + group["Cost Transaction Currency"].sum()
                        material_quantity_by_currency[currency] = material_quantity_by_currency.get(currency, 0) + group["Quantity"].sum()

            for currency, cost in material_cost_by_currency.items():
                cost_data.append({
                    "Project Number": project_number,
                    "Base Material": material,
                    "Product Code": ", ".join(codes),
                    "Transaction Currency": currency,
                    "Cost": cost,
                    "SIZE (inch)": material_info[material]["SIZE (inch)"],
                    "Quantity": material_quantity_by_currency[currency],
                    "Average Size (inch)": material_info[material]["Average Size (inch)"],
                    "Weight": material_info[material]["Weight"],
                    "Average Weight": material_info[material]["Average Weight"]
                })

        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Valve"
        cost_df = pd.DataFrame(cost_data)
        output_file = os.path.join(result_folder_path,
                                   f"MP{project_number}_Valve_MatCurrency_CostAnalyze_{timestamp}.xlsx")
        cost_df.to_excel(output_file, index=False)

        ExportReportsGraphics.plot_valve_cost_quantity_per_currency(cost_df, project_number)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")

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

    required_columns = ["Pipe Base Material", "Product Code", "Total QTY to commit", "Qty confirmed in design", "SIZE", "SBM scope", "Tag Number"]
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
            tag_number = material_rows["Tag Number"]
            total_qty_design = material_rows["Qty confirmed in design"].sum()

            material_info[material] = {
                "Total QTY to commit": total_qty_to_commit,
                'Tag Number': tag_number,
                "Qty confirmed in design": total_qty_design
            }

        ExportReportsGraphics.plot_bolt_quantity_difference(material_info, project_number)

        material_cost_analyze_bolt(project_number, material_codes, material_info)
        material_currency_cost_analyze_bolt(project_number, material_codes, material_info)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")


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
    unmatched_data = []

    if "Product Code" in df.columns and "Cost Project Currency" in df.columns and "Quantity" in df.columns:
        for material, codes in material_codes.items():
            material_cost = 0
            material_quantity = 0

            for code in codes:
                cost_rows = df[df["Product Code"].apply(lambda x: isinstance(x, str) and x.startswith(code))]

                # Check if the code has no match
                if cost_rows.empty:
                    unmatched_data.append({
                        "Project Number": project_number,
                        "Base Material": material,
                        "Product Code": code
                    })
                else:
                    material_cost += cost_rows["Cost Project Currency"].sum()
                    material_quantity += cost_rows["Quantity"].sum()

            if material_cost > 0 or material_quantity > 0:
                cost_data.append({
                    "Project Number": project_number,
                    "Base Material": material,
                    "Product Code": ", ".join(codes),
                    "Cost": material_cost,
                    "Currency": "USD",
                    "Quantity from POs": material_quantity,
                    "Total QTY to commit": material_info[material]["Total QTY to commit"],
                    "Qty confirmed in design": material_info[material]["Qty confirmed in design"],
                })

        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Bolt"
        cost_df = pd.DataFrame(cost_data)
        output_file = os.path.join(result_folder_path,
                                   f"MP{project_number}_Bolt_MaterialBased_Cost_Analyze_{timestamp}.xlsx")
        cost_df.to_excel(output_file, index=False)

        # Save the unmatched data to a new Excel file
        if unmatched_data:
            timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
            unmatched_df = pd.DataFrame(unmatched_data)
            output_file_unmatched = os.path.join(result_folder_path, f"Bolt_NotMatched_ProductCode_{timestamp}.xlsx")
            unmatched_df.to_excel(output_file_unmatched, index=False)

        ExportReportsGraphics.plot_bolt_material_cost(cost_df, project_number)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")


def material_currency_cost_analyze_bolt(project_number, material_codes, material_info):
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

    required_columns = ["Product Code", "Quantity", "Cost Transaction Currency", "Transaction Currency"]
    if all(column in df.columns for column in required_columns):
        for material, codes in material_codes.items():
            material_cost_by_currency = {}
            material_quantity_by_currency = {}

            for code in codes:
                cost_rows = df[df["Product Code"].apply(lambda x: isinstance(x, str) and x.startswith(code))]

                if cost_rows.empty:
                    unmatched_data.append({
                        "Project Number": project_number,
                        "Base Material": material,
                        "Product Code": code
                    })
                else:
                    currency_groups = cost_rows.groupby("Transaction Currency")
                    for currency, group in currency_groups:
                        material_cost_by_currency[currency] = material_cost_by_currency.get(currency, 0) + group["Cost Transaction Currency"].sum()
                        material_quantity_by_currency[currency] = material_quantity_by_currency.get(currency, 0) + group["Quantity"].sum()

            for currency, cost in material_cost_by_currency.items():
                cost_data.append({
                    "Project Number": project_number,
                    "Base Material": material,
                    "Product Code": ", ".join(codes),
                    "Transaction Currency": currency,
                    "Cost": cost,
                    "Total QTY to commit": material_info[material]["Total QTY to commit"],
                    "Qty confirmed in design": material_info[material]["Qty confirmed in design"],
                })

    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Bolt"
    cost_df = pd.DataFrame(cost_data)
    output_file = os.path.join(result_folder_path,
                               f"MP{project_number}_Bolt_MatCurrency_CostAnalyze_{timestamp}.xlsx")
    cost_df.to_excel(output_file, index=False)

#                       ------------------------ Structure --------------------



