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

    required_columns = ["Pipe Base Material", "Tag Number", "Total QTY to commit", "Unit Weight", "Total NET weight", "Quantity UOM", "Unit Weight UOM", "Material", "Service Description"]
    if all(col in df.columns for col in required_columns):
        materials = df["Pipe Base Material"].unique()

        for material in materials:
            pipe_base_materials.add(material)
            material_rows = df[df["Pipe Base Material"] == material]

            for _, row in material_rows.iterrows():
                tag_number = row["Tag Number"]

                # Update material_info dictionary for the tag_number
                if tag_number in material_info:
                    # Accumulate quantity and total weight for existing tag_number
                    material_info[tag_number]["Total QTY to commit"] += row["Total QTY to commit"]
                    material_info[tag_number]["Total NET weight"] += row["Total NET weight"]
                else:
                    # Add new entry for tag_number
                    material_info[tag_number] = {
                        "Pipe Base Material": row["Pipe Base Material"],
                        "Material": row["Material"],
                        "Service Description": row["Service Description"],
                        "Total QTY to commit": row["Total QTY to commit"],
                        "Unit Weight": row["Unit Weight"],
                        "Total NET weight": row["Total NET weight"],
                        "Quantity UOM": row["Quantity UOM"],
                        "Unit Weight UOM": row["Unit Weight UOM"]
                    }

                # Store tag_number directly in the tag_numbers dictionary
                tag_numbers[tag_number] = material

        material_cost_analyze_piping_by_tag(project_number, tag_numbers, material_info)
        material_currency_cost_analyze_piping_by_tag(project_number, tag_numbers, material_info)
    else:
        print("Was not possible to find the necessary fields in the file to do the calculation!")



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

    required_columns = ["Cost Project Currency", "Transaction Date", "PO Description", "Tag Number", "Quantity", "UOM", "Cost Object ID", "PO Line Number"]
    if all(column in df.columns for column in required_columns):
        for tag, material in tag_numbers.items():
            # Regular Tags
            regular_rows = df[df["Tag Number"] == tag]

            # Surplus Tags
            surplus_rows = df[
                df["Tag Number"].str.startswith(tag) &
                (df["Tag Number"].str.endswith("-SURPLUS") | df["Tag Number"].str.endswith("-SURPLUS+"))
                ]

            # Combine Regular and Surplus Rows
            combined_rows = pd.concat([regular_rows, surplus_rows])

            if combined_rows.empty:
                tag_info = material_info[tag]
                unmatched_data.append({
                    "Project Number": project_number,
                    "Base Material": material,
                    "Tag Number": tag,
                    "Material": tag_info["Material"],
                    "Service Description": tag_info["Service Description"],
                })
            else:
                for uom in combined_rows["UOM"].unique():
                    uom_rows = combined_rows[combined_rows["UOM"] == uom]
                    tag_info = material_info[tag]
                    total_cost = uom_rows["Cost Project Currency"].sum()
                    material_quantity = uom_rows["Quantity"].sum()

                    po_tr_date = uom_rows["Transaction Date"].unique()
                    po_tr_date = ', '.join(map(str, po_tr_date))
                    po_desc = uom_rows["PO Description"].unique()
                    po_desc = ', '.join(map(str, po_desc))
                    po_number = regular_rows["Cost Object ID"].unique()  # Get PO number for only regular rows
                    po_number = ', '.join(map(str, po_number))
                    po_linenr = regular_rows["PO Line Number"].unique()  # Get PO Line number for only regular rows
                    po_linenr = ', '.join(map(str, po_linenr))

                    calc_weight = abs(material_quantity * tag_info["Unit Weight"])

                    remarks = ""
                    # Add check for UOM mismatch here
                    uom_in_po = uom
                    quantity_uom = tag_info["Quantity UOM"]
                    if quantity_uom != uom_in_po:
                        if not ((quantity_uom == 'PCE' and uom_in_po == 'PCS') or (
                                quantity_uom == 'METER' and uom_in_po == 'm')):
                            remarks += 'Please note difference between MTO and PO UOM\n'

                    surplus_tags = [row["Tag Number"] for _, row in surplus_rows.iterrows()]
                    for surplus_tag in surplus_tags:
                        remarks += f"This tag number has the surplus reference {surplus_tag}\n"

                    if total_cost > 0 or material_quantity > 0:
                        cost_data.append({
                            "Project Number": project_number,
                            "Tag Number": tag,
                            "PO Number": po_number,
                            #"PO Line Number": po_linenr,
                            "Base Material": material,
                            "Cost": total_cost,
                            "Currency": "USD",
                            "Transaction Date": po_tr_date,
                            "PO Description": po_desc,
                            "Total QTY to commit": tag_info["Total QTY to commit"],
                            "Quantity UOM": tag_info["Quantity UOM"],
                            "Unit Weight": tag_info["Unit Weight"],
                            "Total NET weight": tag_info["Total NET weight"],
                            "Weight UOM": tag_info["Unit Weight UOM"],
                            "Quantity in PO": material_quantity,
                            "UOM in PO": uom,
                            "Total Weight using PO quantity": calc_weight,
                            "Remarks": remarks
                        })

                    if not surplus_rows.empty:
                        for _, surplus_row in surplus_rows.iterrows():
                            surplus_quantity = surplus_row["Quantity"]
                            surplus_cost = surplus_row["Cost Project Currency"]
                            surplus_tag = surplus_row["Tag Number"]
                            surplus_po_number = surplus_row["Cost Object ID"]  # Get unique PO number for surplus

                            cost_data.append({
                                "Project Number": project_number,
                                "Tag Number": surplus_tag,
                                "PO Number": surplus_po_number,  # Assign unique PO number to surplus tag
                                #"PO Line Number": surplus_row["PO Line Number"],
                                "Base Material": material,
                                "Cost": surplus_cost,
                                "Currency": "USD",
                                "Transaction Date": po_tr_date,
                                "PO Description": po_desc,
                                "Total QTY to commit": tag_info["Total QTY to commit"],
                                "Quantity UOM": tag_info["Quantity UOM"],
                                "Unit Weight": tag_info["Unit Weight"],
                                "Total NET weight": tag_info["Total NET weight"],
                                "Weight UOM": tag_info["Unit Weight UOM"],
                                "Quantity in PO": surplus_quantity,
                                "UOM in PO": uom,
                                "Total Weight using PO quantity": calc_weight,
                                "Remarks": ""  # Clear remarks for surplus rows
                            })

        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Piping"
        cost_df = pd.DataFrame(cost_data)

        # Calculate column totals
        total_row = {
            "Base Material": "Total Amount:",
            "Cost": cost_df["Cost"].sum(),
            "Currency": "USD"
        }

        # Append total row to cost_data
        cost_data.append(total_row)
        cost_df = pd.DataFrame(cost_data)
        output_file = os.path.join(result_folder_path, f"MP{project_number}_Piping_TagBased_Cost_Analyze_{timestamp}.xlsx")
        cost_df.to_excel(output_file, index=False)

        # Save the unmatched data to a new Excel file
        if unmatched_data:
            timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
            unmatched_df = pd.DataFrame(unmatched_data)
            output_file_unmatched = os.path.join(result_folder_path, f"Piping_NotMatched_TagNumber_{timestamp}.xlsx")
            unmatched_df.to_excel(output_file_unmatched, index=False)

        ExportReportsGraphics.plot_piping_cost_per_po(cost_df, project_number)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")



def material_currency_cost_analyze_piping_by_tag(project_number, tag_numbers, material_info):
    folder_path = "../Data Pool/Ecosys API Data/PO Lines"

    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if str(project_number) in f]

    if not matching_files:
        raise FileNotFoundError(f"No files containing the project number '{project_number}' were found.")

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    cost_data = []
    required_columns = ["Cost Project Currency", "Cost Transaction Currency", "Transaction Date", "PO Description", "Tag Number", "Quantity", "UOM", "Cost Object ID"]

    if all(column in df.columns for column in required_columns):
        for tag, material in tag_numbers.items():
            # Regular Tags
            regular_rows = df[df["Tag Number"] == tag]

            # Surplus Tags
            surplus_rows = df[
                df["Tag Number"].str.startswith(tag) &
                (df["Tag Number"].str.endswith("-SURPLUS") | df["Tag Number"].str.endswith("-SURPLUS+"))
            ]

            # Combine Regular and Surplus Rows
            combined_rows = pd.concat([regular_rows, surplus_rows])

            if combined_rows.empty:
                continue
            else:
                tag_info = material_info[tag]
                uom_groups = combined_rows.groupby("UOM")
                for uom, group in uom_groups:
                    currency_groups = group.groupby("Transaction Currency")
                    for currency, group in currency_groups:
                        key = (currency, uom)
                        tag_quantity_by_currency_uom = tag_info.setdefault("Quantity by Currency and UOM", {})
                        tag_quantity_by_currency_uom[key] = tag_quantity_by_currency_uom.get(key, 0) + group[
                            "Quantity"].sum()

                        po_tr_date = group["Transaction Date"].unique()
                        po_tr_date = ', '.join(map(str, po_tr_date))
                        po_desc = group["PO Description"].unique()
                        po_desc = ', '.join(map(str, po_desc))
                        po_number = group["Cost Object ID"].unique()
                        po_number = ', '.join(map(str, po_number))

                        calc_weight = abs((tag_info["Quantity by Currency and UOM"][(currency, uom)]) * (tag_info["Unit Weight"]))

                        remarks = ""
                        # Add check for UOM mismatch here
                        uom_in_po = uom
                        quantity_uom = tag_info["Quantity UOM"]
                        if quantity_uom != uom_in_po:
                            if not ((quantity_uom == 'PCE' and uom_in_po == 'PCS') or (
                                    quantity_uom == 'METER' and uom_in_po == 'm')):
                                remarks += 'Please note difference between MTO and PO UOM\n'

                        surplus_rows = group[
                            group["Tag Number"].str.contains("-SURPLUS") | group["Tag Number"].str.contains(
                                "-SURPLUS+S")
                        ]

                        if not surplus_rows.empty:
                            for _, surplus_row in surplus_rows.iterrows():
                                surplus_quantity = surplus_row["Quantity"]
                                surplus_cost = surplus_row["Cost Transaction Currency"]
                                surplus_tag = surplus_row["Tag Number"]
                                remarks += f"A surplus item with Tag Number {surplus_tag} found with the Quantity of {surplus_quantity} and with the cost {surplus_cost}\n"

                        cost_data.append({
                            "Project Number": project_number,
                            "Tag Number": tag,
                            "PO Number": po_number,
                            "Base Material": material,
                            "PO Cost": group["Cost Transaction Currency"].sum(),
                            "Transaction Currency": currency,
                            "Project Currency Cost": group["Cost Project Currency"].sum(),
                            "Currency": "USD",
                            "PO Description": po_desc,
                            "Transaction Date": po_tr_date,
                            "Total QTY to commit": tag_info["Total QTY to commit"],
                            "Quantity UOM": tag_info["Quantity UOM"],
                            "Unit Weight": tag_info["Unit Weight"],
                            "Total NET weight": tag_info["Total NET weight"],
                            "Weight UOM": tag_info["Unit Weight UOM"],
                            "Quantity in PO": tag_info["Quantity by Currency and UOM"][(currency, uom)],
                            "UOM in PO": uom,
                            "Total Weight using PO quantity": calc_weight,
                            "Remarks": remarks
                        })

        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Piping"
        cost_df = pd.DataFrame(cost_data)

        output_file = os.path.join(result_folder_path, f"MP{project_number}_Piping_TagCurrency_Cost_Analyze_{timestamp}.xlsx")
        cost_df.to_excel(output_file, index=False)

        # Plot totals
        ExportReportsGraphics.plot_piping_totals(cost_df, project_number)
        ExportReportsGraphics.plot_piping_cost_per_weight_and_totals(cost_df, project_number)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")

# ----------------------------------------- SPC Piping -------------------------


def extract_distinct_tag_numbers_special_piping(folder_path, project_number, material_type):
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

    tag_numbers = {}
    tag_info = {}

    required_columns = ["TagNumber", "Description", "Size (Inch)", "Service", "Weight", "PO Number", "Qty"]
    if all(col in df.columns for col in required_columns):
        tag_numbers = df["TagNumber"].unique()

        for tag_number in tag_numbers:
            tag_rows = df[df["TagNumber"] == tag_number]

            for _, row in tag_rows.iterrows():
                # Update tag_info dictionary for the tag_number
                if tag_number in tag_info:
                    # Accumulate quantity for existing tag_number
                    tag_info[tag_number]["Qty"] += row["Qty"]
                else:
                    # Add new entry for tag_number
                    tag_info[tag_number] = {
                        "Description": row["Description"],
                        "Size (Inch)": row["Size (Inch)"],
                        "Service": row["Service"],
                        "Weight": row["Weight"],
                        "MTO PO Number": row["PO Number"],
                        "Qty": row["Qty"]
                    }

        material_cost_analyze_special_piping_by_tag(project_number, tag_numbers, tag_info)
        material_currency_cost_analyze_special_piping_by_tag(project_number, tag_numbers, tag_info)
    else:
        print("Was not possible to find the necessary fields in the file to do the calculation!")


def material_cost_analyze_special_piping_by_tag(project_number, tag_numbers, tag_info):
    folder_path = "../Data Pool/Ecosys API Data/PO Lines"

    excel_files = [
        f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    matching_files = [f for f in excel_files if str(project_number) in f ]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found."
        )

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    cost_data = []
    unmatched_data = []

    required_columns = ["Cost Project Currency", "Transaction Date", "PO Description", "Tag Number", "Quantity", "UOM", "Cost Object ID", "PO Line Number"]
    if all(column in df.columns for column in required_columns):
        for tag in tag_numbers:
            rows = df[df["Tag Number"] == tag]

            if rows.empty:
                tag_detail = tag_info[tag]
                unmatched_data.append({
                    "Project Number": project_number,
                    "Tag Number": tag,
                    "Description": tag_detail["Description"],
                    "Size (Inch)": tag_detail["Size (Inch)"],
                    "Service": tag_detail["Service"],
                    "Weight": tag_detail["Weight"],
                    "MTO PO Number": tag_detail["MTO PO Number"],
                    "Qty": tag_detail["Qty"]
                })
            else:
                for uom in rows["UOM"].unique():
                    uom_rows = rows[rows["UOM"] == uom]
                    tag_detail = tag_info[tag]
                    total_cost = uom_rows["Cost Project Currency"].sum()
                    quantity_in_po = uom_rows["Quantity"].sum()

                    po_tr_date = uom_rows["Transaction Date"].unique()
                    po_tr_date = ', '.join(map(str, po_tr_date))
                    po_desc = uom_rows["PO Description"].unique()
                    po_desc = ', '.join(map(str, po_desc))
                    po_number = uom_rows["Cost Object ID"].unique()
                    po_number = ', '.join(map(str, po_number))
                    po_linenr = uom_rows["PO Line Number"].unique()
                    po_linenr = ', '.join(map(str, po_linenr))
                    po_revision = uom_rows["PO Revision"].unique()
                    po_revision = ', '.join(map(str, po_revision))

                    remarks = ""

                    cost_data.append({
                        "Project Number": project_number,
                        "Tag Number": tag,
                        "MTO PO Number": tag_detail["MTO PO Number"],
                        "PO Number": po_number,
                        "Description": tag_detail["Description"],
                        "Size (Inch)": tag_detail["Size (Inch)"],
                        "Service": tag_detail["Service"],
                        "MTO Quantity": tag_detail["Qty"],
                        "MTO Weight": tag_detail["Weight"],
                        "Cost": total_cost,
                        "Currency": "USD",
                        "Transaction Date": po_tr_date,
                        "PO Description": po_desc,
                        "PO Revision": po_revision,
                        "Quantity in PO": quantity_in_po,
                        "UOM in PO": uom,
                        "Remarks": remarks
                    })

        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/SPC Piping"
        cost_df = pd.DataFrame(cost_data)

        # Calculate column totals
        total_row = {
            "Description": "Total Amount:",
            "Cost": cost_df["Cost"].sum(),
            "Currency": "USD"
        }

        # Append total row to cost_data
        cost_data.append(total_row)
        cost_df = pd.DataFrame(cost_data)
        output_file = os.path.join(result_folder_path, f"MP{project_number}_SPCPiping_TagBased_Cost_{timestamp}.xlsx")
        cost_df.to_excel(output_file, index=False)

        # Save the unmatched data to a new Excel file
        if unmatched_data:
            timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
            unmatched_df = pd.DataFrame(unmatched_data)
            output_file_unmatched = os.path.join(result_folder_path, f"Special_Piping_NotMatched_TagNumber_{timestamp}.xlsx")
            unmatched_df.to_excel(output_file_unmatched, index=False)

        # Plot totals
        ExportReportsGraphics.plot_special_piping_cost_per_po_and_supplier(cost_df, project_number)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")


def material_currency_cost_analyze_special_piping_by_tag(project_number, tag_numbers, material_info):
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
    required_columns = ["Cost Project Currency", "Transaction Date", "PO Description", "Tag Number", "Quantity", "UOM", "Cost Object ID", "Cost Transaction Currency"]

    if all(column in df.columns for column in required_columns):
        for tag in tag_numbers:
            rows = df[df["Tag Number"] == tag]

            if rows.empty:
                continue
            else:
                tag_info = material_info[tag]
                uom_groups = rows.groupby("UOM")
                for uom, group in uom_groups:
                    currency_groups = group.groupby("Transaction Currency")
                    for currency, group in currency_groups:
                        key = (currency, uom)
                        tag_quantity_by_currency_uom = tag_info.setdefault("Quantity by Currency and UOM", {})
                        tag_quantity_by_currency_uom[key] = tag_quantity_by_currency_uom.get(key, 0) + group[
                            "Quantity"].sum()

                        po_tr_date = group["Transaction Date"].unique()
                        po_tr_date = ', '.join(map(str, po_tr_date))
                        po_desc = group["PO Description"].unique()
                        po_desc = ', '.join(map(str, po_desc))
                        po_number = group["Cost Object ID"].unique()
                        po_number = ', '.join(map(str, po_number))
                        po_revision = group["PO Revision"].unique()
                        po_revision = ', '.join(map(str, po_revision))

                        calc_weight = abs((tag_info["Quantity by Currency and UOM"][(currency, uom)]) * (tag_info["Weight"]))

                        cost_data.append({
                            "Project Number": project_number,
                            "Tag Number": tag,
                            "MTO PO Number": tag_info["MTO PO Number"],
                            "PO Number": po_number,
                            "PO Cost": group["Cost Transaction Currency"].sum(),
                            "Transaction Currency": currency,
                            "Project Currency Cost": group["Cost Project Currency"].sum(),
                            "Currency": "USD",
                            "PO Description": po_desc,
                            "Transaction Date": po_tr_date,
                            "PO Revision": po_revision,
                            "MTO Weight": tag_info["Weight"],
                            "MTO Quantity": tag_info["Qty"],
                            "Quantity in PO": tag_info["Quantity by Currency and UOM"][(currency, uom)],
                            "UOM in PO": uom,
                            "Total Weight using PO quantity": calc_weight,
                        })

        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/SPC Piping"
        cost_df = pd.DataFrame(cost_data)

        output_file = os.path.join(result_folder_path, f"MP{project_number}_SPCPiping_TagCurrency_Cost_{timestamp}.xlsx")
        cost_df.to_excel(output_file, index=False)

        # Plot totals
        ExportReportsGraphics.plot_special_piping_cost_per_weight_and_totals(cost_df, project_number)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")

