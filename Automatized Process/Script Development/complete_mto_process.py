import os
import numpy as np
import pandas as pd
import ExportPDFreports


folder_path_unity = "../Data Pool/Data Hub Materials/MP17033 UNITY"
folder_path_prosperity = "../Data Pool/Data Hub Materials/MP17043 PROSPERITY"


#Complete analyze of those different MTO files - Material type, Weight, Quantities
def complete_mto_data_analyze(project_number):

    if project_number == "17033":
        print("Unity Complete MTO analyze on going!")
        #Call each function and collect their data
        piping_data, piping_sbm_data, piping_data_yard, total_qty_commit_pieces, total_qty_commit_m, total_piping_net_weight, total_piping_sbm_net_weight, total_piping_yard_net_weight, total_qty_commit_pieces_sbm, total_qty_commit_pieces_yard, total_qty_commit_m_sbm, total_qty_commit_m_yard = get_piping_mto_data(project_number)
        valve_data, valve_data_sbm, valve_data_yard, total_valve_weight, total_sbm_valve_weight, total_yard_valve_weight = get_valve_mto_data(project_number)
        bolt_data_total_qty_commit, bolt_sbm_data_total_qty_commit, bolt_yard_data_total_qty_commit = get_bolt_mto_data(project_number)
        structure_totals_m2, structure_totals_m, structure_totals_pcs, structure_total_gross_weight, structure_total_wastage, structure_total_qty_pcs, structure_total_qty_m2, structure_total_qty_m = get_structure_mto_data(project_number)
        total_spcpip_data_qty, total_spcpip_data_weight, total_spcpip_sbm_data_qty, total_spcpip_sbm_data_weight, total_spcpip_yard_data_qty, total_spcpip_yard_data_weight = get_specialpip_mto_data(project_number)

        #Piping Extra Data details
        total_matched_tags_pip, total_unmatched_tags_pip, total_surplus_tags_pip, total_weight_pip, total_quantity_by_uom_pip, overall_cost_pip, total_cost_by_material_pip, unique_cost_object_ids_pip, total_surplus_cost_pip, unique_surplus_cost_object_ids_pip, total_surplus_plus_tags, total_po_quantity_piece, total_po_quantity_meter = get_piping_extra_details(project_number, "Piping")
        #Valve Extra Data details
        total_quantity_vlv, overall_cost_vlv, cost_by_general_description_vlv = get_valve_extra_details(project_number, "Valve")
        #Bolt Extra Data details
        total_po_quantity_blt, overall_cost_blt, cost_by_pipe_base_material_blt, missing_product_codes_blt = get_bolt_extra_details(project_number, "Bolt")
        #Special Piping Extra details
        total_matched_tags_spc, total_unmatched_tags_spc, total_quantity_by_uom_spc, total_cost_spc, po_list_spc = get_spc_piping_extra_details(project_number, "Piping")

        #Get PO Header overall amount
        project_total_cost_and_hours = get_project_total_cost_hours(project_number)

        # Pass the data frame to export functions
        ExportPDFreports.generate_unity_complete_analyze_process_pdf(piping_data, piping_sbm_data, piping_data_yard, total_qty_commit_pieces, total_qty_commit_m, total_piping_net_weight, total_piping_sbm_net_weight, total_piping_yard_net_weight, total_qty_commit_pieces_sbm, total_qty_commit_pieces_yard, total_qty_commit_m_sbm, total_qty_commit_m_yard, valve_data_sbm, valve_data_yard, total_valve_weight, total_sbm_valve_weight, total_yard_valve_weight, bolt_data_total_qty_commit, bolt_sbm_data_total_qty_commit, bolt_yard_data_total_qty_commit, structure_totals_m2, structure_totals_m, structure_totals_pcs,
                                                               total_matched_tags_pip, total_unmatched_tags_pip, total_surplus_tags_pip, total_weight_pip, total_quantity_by_uom_pip, overall_cost_pip, total_cost_by_material_pip, unique_cost_object_ids_pip, total_surplus_cost_pip, unique_surplus_cost_object_ids_pip, total_spcpip_data_weight, total_spcpip_data_qty, total_spcpip_sbm_data_weight, total_spcpip_sbm_data_qty, total_spcpip_yard_data_qty, total_spcpip_yard_data_weight,
                                                               total_quantity_vlv, overall_cost_vlv, cost_by_general_description_vlv, total_po_quantity_blt, overall_cost_blt, cost_by_pipe_base_material_blt, missing_product_codes_blt, structure_total_gross_weight, structure_total_wastage, structure_total_qty_pcs, structure_total_qty_m2, structure_total_qty_m, total_matched_tags_spc, total_unmatched_tags_spc, total_quantity_by_uom_spc, total_cost_spc, po_list_spc, project_total_cost_and_hours, total_surplus_plus_tags, total_po_quantity_piece, total_po_quantity_meter)
    elif project_number == "17043":
        print("Prosperity Complete MTO analyze on going!")
        piping_data, piping_sbm_data, piping_data_yard, total_qty_commit_pieces, total_qty_commit_m, total_piping_net_weight, total_piping_sbm_net_weight, total_piping_yard_net_weight, total_qty_commit_pieces_sbm, total_qty_commit_pieces_yard, total_qty_commit_m_sbm, total_qty_commit_m_yard = get_piping_mto_data(project_number)
        total_valve_weight, total_sbm_valve_weight, total_yard_valve_weight = get_valve_mto_data(project_number)
        bolt_data_total_net_weight, bolt_data_total_qty_to_complete, bolt_data_total_to_complete_weight, bolt_data_total_qty_commit, bolt_sbm_data_total_qty_commit, bolt_sbm_total_net_weight, bolt_sbm_total_qty_to_complete, bolt_sbm_total_to_complete_weight, bolt_yard_data_total_qty_commit, bolt_yard_total_net_weight, bolt_yard_total_qty_to_complete, bolt_yard_total_to_complete_weight = get_bolt_mto_data(project_number)
        structure_totals_m2, structure_totals_m, structure_totals_pcs, structure_total_gross_weight, structure_total_net_weight, structure_total_wastage, structure_total_qty_pcs, structure_total_req_pcs, structure_total_qty_m2, structure_total_req_m2, structure_total_qty_m, structure_total_req_m = get_structure_mto_data(project_number)
        total_spcpip_data_qty, total_spcpip_data_weight, total_spcpip_sbm_data_qty, total_spcpip_sbm_data_weight, total_spcpip_yard_data_qty, total_spcpip_yard_data_weight = get_specialpip_mto_data(project_number)

        # Get PO Header overall amount
        project_total_cost_and_hours = get_project_total_cost_hours(project_number)


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

    #Unity
    if project_number == "17033":
        excel_files = [f for f in os.listdir(folder_path_unity) if f.endswith(".xlsx") or f.endswith(".xls")]
        matching_files = [f for f in excel_files if material_type in f]

        if not matching_files:
            raise FileNotFoundError(
                f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

        most_recent_file = get_most_recent_file(folder_path_unity, matching_files)
        file_path = os.path.join(folder_path_unity, most_recent_file)
        df = pd.read_excel(file_path)

        columns_to_extract = ["Project Number", "Pipe Base Material", "SBM scope", "Total QTY to commit", "Quantity UOM",
                              "Unit Weight", "Unit Weight UOM", "Total NET weight"]

        if df["Project Number"].astype(str).str.contains(str(project_number)).any():
            extract_columns = [column for column in df.columns if any(col in str(column) for col in columns_to_extract)]
            filtered_df = df[extract_columns]
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

            total_qty_commit_pieces = piping_data[piping_data['Quantity UOM'] == "PCE"]['Total QTY to commit'].sum()
            total_qty_commit_m = piping_data[piping_data['Quantity UOM'] == "METER"]['Total QTY to commit'].sum()
            total_qty_commit_pieces_sbm = piping_sbm_data[piping_sbm_data['Quantity UOM'] == "PCE"]['Total QTY to commit'].sum()
            total_qty_commit_pieces_yard = piping_data_yard[piping_data_yard['Quantity UOM'] == "PCE"]['Total QTY to commit'].sum()
            total_qty_commit_m_sbm = piping_sbm_data[piping_sbm_data['Quantity UOM'] == "METER"]['Total QTY to commit'].sum()
            total_qty_commit_m_yard = piping_data_yard[piping_data_yard['Quantity UOM'] == "METER"]['Total QTY to commit'].sum()
            total_piping_net_weight = piping_data['Total NET weight'].sum()
            total_piping_sbm_net_weight = piping_sbm_data['Total NET weight'].sum()
            total_piping_yard_net_weight = piping_data_yard['Total NET weight'].sum()

            return piping_data, piping_sbm_data, piping_data_yard, total_qty_commit_pieces, total_qty_commit_m, total_piping_net_weight, total_piping_sbm_net_weight, total_piping_yard_net_weight, total_qty_commit_pieces_sbm, total_qty_commit_pieces_yard, total_qty_commit_m_sbm, total_qty_commit_m_yard

        else:
            raise ValueError(f"No data found for project number '{project_number}'.")
    #PROSPERITY
    elif project_number == "17043":
        excel_files = [f for f in os.listdir(folder_path_prosperity) if f.endswith(".xlsx") or f.endswith(".xls")]
        matching_files = [f for f in excel_files if material_type in f]

        if not matching_files:
            raise FileNotFoundError(
                f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

        most_recent_file = get_most_recent_file(folder_path_prosperity, matching_files)
        file_path = os.path.join(folder_path_prosperity, most_recent_file)
        df = pd.read_excel(file_path)

        columns_to_extract = ["Project Number", "Pipe Base Material", "SBM scope", "Total QTY to commit",
                              "Quantity UOM", "Unit Weight", "Total NET weight"]

        if df["Project Number"].astype(str).str.contains(str(project_number)).any():
            extract_columns = [column for column in df.columns if any(col in str(column) for col in columns_to_extract)]
            filtered_df = df[extract_columns]
            extract_df_sbm = filtered_df[(filtered_df['SBM scope'] == "1") & (filtered_df['SBM scope'].notnull())]
            extract_df_yard = filtered_df[(filtered_df['SBM scope'] == "0") & (filtered_df['SBM scope'].notnull())]

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

            total_qty_commit_pieces = piping_data[piping_data['Quantity UOM'] == "PCE"]['Total QTY to commit'].sum()
            total_qty_commit_m = piping_data[piping_data['Quantity UOM'] == "METER"]['Total QTY to commit'].sum()
            total_qty_commit_pieces_sbm = piping_sbm_data[piping_sbm_data['Quantity UOM'] == "PCE"]['Total QTY to commit'].sum()
            total_qty_commit_pieces_yard = piping_data_yard[piping_data_yard['Quantity UOM'] == "PCE"]['Total QTY to commit'].sum()
            total_qty_commit_m_sbm = piping_sbm_data[piping_sbm_data['Quantity UOM'] == "METER"]['Total QTY to commit'].sum()
            total_qty_commit_m_yard = piping_data_yard[piping_data_yard['Quantity UOM'] == "METER"]['Total QTY to commit'].sum()
            total_piping_net_weight = piping_data['Total NET weight'].sum()
            total_piping_sbm_net_weight = piping_sbm_data['Total NET weight'].sum()
            total_piping_yard_net_weight = piping_data_yard['Total NET weight'].sum()

            return piping_data, piping_sbm_data, piping_data_yard, total_qty_commit_pieces, total_qty_commit_m, total_piping_net_weight, total_piping_sbm_net_weight, total_piping_yard_net_weight, total_qty_commit_pieces_sbm, total_qty_commit_pieces_yard, total_qty_commit_m_sbm, total_qty_commit_m_yard

        else:
            raise ValueError(f"No data found for project number '{project_number}'.")


#VALVE MTO DATA
def get_valve_mto_data(project_number):
    material_type = "Valve"

    #UNITY
    if project_number == "17033":
        excel_files = [f for f in os.listdir(folder_path_unity) if f.endswith(".xlsx") or f.endswith(".xls")]
        matching_files = [f for f in excel_files if material_type in f]

        if not matching_files:
            raise FileNotFoundError(
                f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

        most_recent_file = get_most_recent_file(folder_path_unity, matching_files)
        file_path = os.path.join(folder_path_unity, most_recent_file)
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

            # Calculate the sum of the total valve weight for all three categories
            total_valve_weight = valve_data['Weight'].sum()
            total_sbm_valve_weight = valve_data_sbm['Weight'].sum()
            total_yard_valve_weight = valve_data_yard['Weight'].sum()

            return valve_data, valve_data_sbm, valve_data_yard, total_valve_weight, total_sbm_valve_weight, total_yard_valve_weight
        else:
            raise ValueError(f"No data found for project number '{project_number}'.")
    #PROSPERITY
    elif project_number == "17043":
        excel_files = [f for f in os.listdir(folder_path_prosperity) if f.endswith(".xlsx") or f.endswith(".xls")]
        matching_files = [f for f in excel_files if material_type in f]

        if not matching_files:
            raise FileNotFoundError(
                f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

        most_recent_file = get_most_recent_file(folder_path_prosperity, matching_files)
        file_path = os.path.join(folder_path_prosperity, most_recent_file)
        df = pd.read_excel(file_path)

        # Set the columns to extract
        columns_to_extract = ["Tag Status", "Project Number", "General Material Description", "Quantity", "Dry Weight [kg]",
                              "Remarks", "Valve Size", "Scope Of Supply"]

        # Check if the project number matches the specified one
        if df["Scope Of Supply"].astype(str).str.contains("SBM").any():
            # Find the specified columns
            extract_columns = []
            for column in df.columns:
                if any(col in str(column) for col in columns_to_extract):
                    extract_columns.append(column)

            # Filter the DataFrame and extract the desired data
            filtered_df = df[extract_columns]

            # Filter based on SBM scope and notnull values
            extract_df_sbm = filtered_df[(filtered_df['Tag Status'] == "New") | (filtered_df['Scope Of Supply'] == "SBM")]
            extract_df_yard = filtered_df[(filtered_df['Tag Status'] == "New") | (filtered_df['Scope Of Supply'] == "YARD")]

            '''valve_data = filtered_df.groupby(["General Material Description"]).agg({
                 "Valve Size": "mean",
                 "Dry Weight [kg]": "sum"
             }).reset_index()

             valve_data_sbm = extract_df_sbm.groupby(["General Material Description"]).agg({
                 "Valve Size": "mean",
                 "Dry Weight [kg]": "sum"
             }).reset_index()

             valve_data_yard = extract_df_yard.groupby(["General Material Description"]).agg({
                 "Valve Size": "mean",
                 "Dry Weight [kg]": "sum"
             }).reset_index() '''

            #Calculate the sum of the total valve weight for all three categories
            total_valve_weight = filtered_df['Dry Weight [kg]'].sum()
            total_sbm_valve_weight = extract_df_sbm['Dry Weight [kg]'].sum()
            total_yard_valve_weight = extract_df_yard['Dry Weight [kg]'].sum()

            return total_valve_weight, total_sbm_valve_weight, total_yard_valve_weight
        else:
            raise ValueError(f"No data found for project number '{project_number}'.")


#BOLT MTO DATA
def get_bolt_mto_data(project_number):
    material_type = "Bolt"

    #UNITY
    if project_number == "17033":
        excel_files = [f for f in os.listdir(folder_path_unity) if f.endswith(".xlsx") or f.endswith(".xls")]
        matching_files = [f for f in excel_files if material_type in f]

        if not matching_files:
            raise FileNotFoundError(
                f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

        most_recent_file = get_most_recent_file(folder_path_unity, matching_files)
        file_path = os.path.join(folder_path_unity, most_recent_file)
        df = pd.read_excel(file_path)

        columns_to_extract = ["Project Number", "Pipe Base Material", "SBM scope", "Total QTY to commit", "Quantity UOM"]

        if df["Project Number"].astype(str).str.contains(str(project_number)).any():
            extract_columns = [column for column in df.columns if any(col in str(column) for col in columns_to_extract)]
            filtered_df = df[extract_columns]
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

            bolt_data_total_qty_commit = bolt_data['Total QTY to commit'].sum()
            bolt_sbm_data_total_qty_commit = bolt_sbm_data['Total QTY to commit'].sum()
            bolt_yard_data_total_qty_commit = bolt_data_yard['Total QTY to commit'].sum()

            return bolt_data_total_qty_commit, bolt_sbm_data_total_qty_commit, bolt_yard_data_total_qty_commit

        else:
            raise ValueError(f"No data found for project number '{project_number}'.")
    #PROSPERITY
    elif project_number == "17043":
        excel_files = [f for f in os.listdir(folder_path_prosperity) if f.endswith(".xlsx") or f.endswith(".xls")]
        matching_files = [f for f in excel_files if material_type in f]

        if not matching_files:
            raise FileNotFoundError(
                f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

        most_recent_file = get_most_recent_file(folder_path_prosperity, matching_files)
        file_path = os.path.join(folder_path_prosperity, most_recent_file)
        df = pd.read_excel(file_path)

        columns_to_extract = ["Project Number", "Pipe Base Material", "SBM scope", "Total QTY to commit",
                              "Quantity UOM", "Total NET weight", "Quantity to complete", "Total Qty to complete weight"]

        if df["Project Number"].astype(str).str.contains(str(project_number)).any():
            extract_columns = [column for column in df.columns if any(col in str(column) for col in columns_to_extract)]
            filtered_df = df[extract_columns]
            extract_df_sbm = filtered_df[(filtered_df['SBM scope'] == True) & (filtered_df['SBM scope'].notnull())]
            extract_df_yard = filtered_df[(filtered_df['SBM scope'] == False) & (filtered_df['SBM scope'].notnull())]

            bolt_data = filtered_df.groupby(["Pipe Base Material"]).agg({
                "Total QTY to commit": "sum",
                "Total NET weight": "sum",
                "Quantity to complete": "sum",
                "Total Qty to complete weight": "sum"
            }).reset_index()

            bolt_sbm_data = extract_df_sbm.groupby(["Pipe Base Material"]).agg({
                "Total QTY to commit": "sum",
                "Total NET weight": "sum",
                "Quantity to complete": "sum",
                "Total Qty to complete weight": "sum"
            }).reset_index()

            bolt_data_yard = extract_df_yard.groupby(["Pipe Base Material"]).agg({
                "Total QTY to commit": "sum",
                "Total NET weight": "sum",
                "Quantity to complete": "sum",
                "Total Qty to complete weight": "sum"
            }).reset_index()

            bolt_data_total_qty_commit = bolt_data['Total QTY to commit'].sum()
            bolt_data_total_net_weight = bolt_data['Total NET weight'].sum()
            bolt_data_total_qty_to_complete = bolt_data['Quantity to complete'].sum()
            bolt_data_total_to_complete_weight = bolt_data['Total Qty to complete weight'].sum()
            bolt_sbm_data_total_qty_commit = bolt_sbm_data['Total QTY to commit'].sum()
            bolt_sbm_total_net_weight = bolt_sbm_data['Total NET weight'].sum()
            bolt_sbm_total_qty_to_complete = bolt_sbm_data['Quantity to complete'].sum()
            bolt_sbm_total_to_complete_weight = bolt_sbm_data['Total Qty to complete weight'].sum()
            bolt_yard_data_total_qty_commit = bolt_data_yard['Total QTY to commit'].sum()
            bolt_yard_total_net_weight = bolt_data_yard['Total NET weight'].sum()
            bolt_yard_total_qty_to_complete = bolt_data_yard['Quantity to complete'].sum()
            bolt_yard_total_to_complete_weight = bolt_data_yard['Total Qty to complete weight'].sum()

            return bolt_data_total_net_weight, bolt_data_total_qty_to_complete, bolt_data_total_to_complete_weight, bolt_data_total_qty_commit, bolt_sbm_data_total_qty_commit, bolt_sbm_total_net_weight, bolt_sbm_total_qty_to_complete, \
                   bolt_sbm_total_to_complete_weight, bolt_yard_data_total_qty_commit, bolt_yard_total_net_weight, bolt_yard_total_qty_to_complete, bolt_yard_total_to_complete_weight
        else:
            raise ValueError(f"No data found for project number '{project_number}'.")


#STRUCTURAL MTO DATA
def get_structure_mto_data(project_number):
    material_type = "Structure"

    #UNITY
    if project_number == "17033":
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

        # Calculate the total values for each group
        structure_total_qty_pcs = structure_totals_pcs['Total QTY to commit'].sum()
        structure_total_qty_m2 = structure_totals_m2['Total QTY to commit'].sum()
        structure_total_qty_m = structure_totals_m['Total QTY to commit'].sum()
        structure_total_gross_weight = structure_totals_pcs['Total Gross Weight'].sum() + structure_totals_m2['Total Gross Weight'].sum() + structure_totals_m['Total Gross Weight'].sum()
        structure_total_wastage = structure_totals_pcs['Wastage Quantity'].sum() + structure_totals_m2['Wastage Quantity'].sum() + structure_totals_m['Wastage Quantity'].sum()

        return structure_totals_m2, structure_totals_m, structure_totals_pcs, structure_total_gross_weight, structure_total_wastage, structure_total_qty_pcs, structure_total_qty_m2, structure_total_qty_m
    # PROSPERITY
    elif project_number == "17043":

        excel_files = [f for f in os.listdir(folder_path_prosperity) if f.endswith(".xlsx") or f.endswith(".xls")]
        matching_files = [f for f in excel_files if str(project_number) in f and material_type in f]

        if not matching_files:
            raise FileNotFoundError(
                f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

        most_recent_file = get_most_recent_file(folder_path_prosperity, matching_files)
        file_path = os.path.join(folder_path_prosperity, most_recent_file)
        df = pd.read_excel(file_path)

        structure_qty_pce = df[(df['Quantity UOM'] == "PCS")]
        structure_qty_m = df[(df['Quantity UOM'] == "m")]
        structure_qty_m2 = df[(df['Quantity UOM'] == "m2")]

        structure_totals_pcs = structure_qty_pce.groupby(["Quantity UOM"]).agg({
            'Total QTY to commit': 'sum',
            'Total NET weight': 'sum',
            'Unit Weight': 'mean',
            'Required Qty': 'sum',
            'Thickness': 'sum',
            'Wastage Quantity': 'sum',
            'Quantity Including Wastage': 'sum',
            'Total Gross Weight': 'sum', }).reset_index()

        structure_totals_m = structure_qty_m.groupby(["Quantity UOM"]).agg({
            'Total QTY to commit': 'sum',
            'Total NET weight': 'sum',
            'Unit Weight': 'mean',
            'Required Qty': 'sum',
            'Thickness': 'sum',
            'Wastage Quantity': 'sum',
            'Quantity Including Wastage': 'sum',
            'Total Gross Weight': 'sum', }).reset_index()

        structure_totals_m2 = structure_qty_m2.groupby(["Quantity UOM"]).agg({
            'Total QTY to commit': 'sum',
            'Total NET weight': 'sum',
            'Unit Weight': 'mean',
            'Required Qty': 'sum',
            'Thickness': 'sum',
            'Wastage Quantity': 'sum',
            'Quantity Including Wastage': 'sum',
            'Total Gross Weight': 'sum', }).reset_index()

        # Calculate the total values for each group
        structure_total_qty_pcs = structure_totals_pcs['Total QTY to commit'].sum()
        structure_total_req_pcs = structure_totals_pcs['Required Qty'].sum()
        structure_total_qty_m2 = structure_totals_m2['Total QTY to commit'].sum()
        structure_total_req_m2 = structure_totals_pcs['Required Qty'].sum()
        structure_total_qty_m = structure_totals_m['Total QTY to commit'].sum()
        structure_total_req_m = structure_totals_pcs['Required Qty'].sum()
        structure_total_gross_weight = structure_totals_pcs['Total Gross Weight'].sum() + structure_totals_m2['Total Gross Weight'].sum() + structure_totals_m['Total Gross Weight'].sum()
        structure_total_net_weight = structure_totals_pcs['Total NET weight'].sum() + structure_totals_m2['Total NET weight'].sum() + structure_totals_m['Total NET weight'].sum()
        structure_total_wastage = structure_totals_pcs['Wastage Quantity'].sum() + structure_totals_m2['Wastage Quantity'].sum() + structure_totals_m['Wastage Quantity'].sum()

        return structure_totals_m2, structure_totals_m, structure_totals_pcs, structure_total_gross_weight, structure_total_net_weight, structure_total_wastage, structure_total_qty_pcs, structure_total_req_pcs, structure_total_qty_m2, structure_total_req_m2, structure_total_qty_m, structure_total_req_m


#SPECIAL PIPING MTO DATA
def get_specialpip_mto_data(project_number):
    material_type = "Special PIP"
    # UNITY
    if project_number == "17033":
        excel_files = [f for f in os.listdir(folder_path_unity) if f.endswith(".xlsx") or f.endswith(".xls")]
        matching_files = [f for f in excel_files if material_type in f]

        if not matching_files:
            raise FileNotFoundError(
                f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

        most_recent_file = get_most_recent_file(folder_path_unity, matching_files)
        file_path = os.path.join(folder_path_unity, most_recent_file)
        df = pd.read_excel(file_path)

        # Set the columns to extract
        columns_to_extract = ["TagNumber", "Size (Inch)", "Weight", "Remarks", "PO Number", "Qty"]

        # Find the specified columns
        extract_columns = []
        for column in df.columns:
            if any(col in str(column) for col in columns_to_extract):
                extract_columns.append(column)

        # Filter the DataFrame and extract the desired data
        filtered_df = df.loc[:, extract_columns].copy()

        # Convert Qty and Weight columns to numeric type
        filtered_df.loc[:, "Qty"] = pd.to_numeric(filtered_df["Qty"], errors="coerce")
        filtered_df.loc[:, "Weight"] = pd.to_numeric(filtered_df["Weight"], errors="coerce")

        # Filter based on SBM scope and notnull values
        extract_df_sbm = filtered_df[(filtered_df['PO Number'] != "BY YARD") & (filtered_df['PO Number'].notnull())]
        extract_df_yard = filtered_df[(filtered_df['PO Number'] == "BY YARD") & (filtered_df['PO Number'].notnull())]

        # Calculate the total quantities and total weights for each category
        total_spcpip_data_qty = filtered_df["Qty"].sum()
        total_spcpip_data_weight = filtered_df["Weight"].sum()
        total_spcpip_sbm_data_qty = extract_df_sbm["Qty"].sum()
        total_spcpip_sbm_data_weight = extract_df_sbm["Weight"].sum()
        total_spcpip_yard_data_qty = extract_df_yard["Qty"].sum()
        total_spcpip_yard_data_weight = extract_df_yard["Weight"].sum()

        return total_spcpip_data_qty, total_spcpip_data_weight, total_spcpip_sbm_data_qty, total_spcpip_sbm_data_weight, total_spcpip_yard_data_qty, total_spcpip_yard_data_weight
    # PROSPERITY
    elif project_number == "17043":
        excel_files = [f for f in os.listdir(folder_path_prosperity) if f.endswith(".xlsx") or f.endswith(".xls")]
        matching_files = [f for f in excel_files if material_type in f]

        if not matching_files:
            raise FileNotFoundError(
                f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

        most_recent_file = get_most_recent_file(folder_path_prosperity, matching_files)
        file_path = os.path.join(folder_path_prosperity, most_recent_file)
        df = pd.read_excel(file_path)

        # Set the columns to extract
        columns_to_extract = ["TAG Number", "Type", "Tag Status", "Scope Of Supply", "Dry Weight [kg]", "PO Number", "Size"]

        # Find the specified columns
        extract_columns = []
        for column in df.columns:
            if any(col in str(column) for col in columns_to_extract):
                extract_columns.append(column)

        # Filter the DataFrame and extract the desired data
        filtered_df = df.loc[:, extract_columns].copy()

        # Convert Qty and Weight columns to numeric type
        filtered_df.loc[:, "Size"] = pd.to_numeric(filtered_df["Size"], errors="coerce")
        filtered_df.loc[:, "Dry Weight [kg]"] = pd.to_numeric(filtered_df["Dry Weight [kg]"], errors="coerce")

        # Filter based on SBM scope and notnull values
        extract_df_sbm = filtered_df[(filtered_df['Scope Of Supply'] == "SBM") & (filtered_df['Type'] == "SpecialPipingItem")]
        extract_df_yard = filtered_df[(filtered_df['Scope Of Supply'] == "YARD") & (filtered_df['Type'] == "SpecialPipingItem")]

        # Calculate the total quantities and total weights for each category
        total_spcpip_data_qty = filtered_df["Size"].sum()
        total_spcpip_data_weight = filtered_df["Dry Weight [kg]"].sum()
        total_spcpip_sbm_data_qty = extract_df_sbm["Size"].sum()
        total_spcpip_sbm_data_weight = extract_df_sbm["Dry Weight [kg]"].sum()
        total_spcpip_yard_data_qty = extract_df_yard["Size"].sum()
        total_spcpip_yard_data_weight = extract_df_yard["Dry Weight [kg]"].sum()

        return total_spcpip_data_qty, total_spcpip_data_weight, total_spcpip_sbm_data_qty, total_spcpip_sbm_data_weight, total_spcpip_yard_data_qty, total_spcpip_yard_data_weight


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
    total_surplus_tags = second_df["Tag Number"].str.contains("-SURPLUS", na=False).sum()
    total_surplus_plus_tags = second_df["Tag Number"].str.contains("SURPLUS+S", na=False).sum()
    total_weight = sum(entry["PO Weight"] for entry in cost_data)
    total_quantity_by_uom = {}

    for entry in cost_data:
        for uom, quantity in entry["PO Quantity by UOM"].items():
            if uom not in total_quantity_by_uom:
                total_quantity_by_uom[uom] = 0.0
            total_quantity_by_uom[uom] += quantity

    total_cost = sum(entry["Total Cost"] for entry in cost_data)

    # Calculate the sum of 'EA', 'PCE', and 'PCS'
    total_po_quantity_piece = total_quantity_by_uom.get('EA', 0) + total_quantity_by_uom.get('PCE', 0) + total_quantity_by_uom.get('PCS', 0)

    # Calculate the sum of 'METER' and 'm'
    total_po_quantity_meter = total_quantity_by_uom.get('METER', 0) + total_quantity_by_uom.get('m', 0)

    # Calculate total cost by material
    total_cost_by_material = {}
    overall_cost_pip = total_cost

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
        overall_cost_pip,
        total_cost_by_material,
        unique_cost_object_ids,
        total_surplus_cost,
        unique_surplus_cost_object_ids,
        total_surplus_plus_tags,
        total_po_quantity_piece,
        total_po_quantity_meter
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
    total_matched_tags_spc = len(cost_data)
    total_unmatched_tags_spc = len(first_tag_numbers) - len(cost_data)
    total_quantity_by_uom_spc = sum(entry["PO Quantity"] for entry in cost_data)
    total_cost_spc = sum(entry["Total Cost"] for entry in cost_data)

    # Get distinct PO Numbers
    po_list_spc = list(set(entry["PO Number"] for entry in cost_data))

    # Return the calculated data along with the unique Cost Object IDs, total surplus cost, and unique surplus Cost Object IDs
    return (
        total_matched_tags_spc,
        total_unmatched_tags_spc,
        total_quantity_by_uom_spc,
        total_cost_spc,
        po_list_spc
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

    overall_cost_bolt = total_cost

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
        overall_cost_bolt,
        cost_by_pipe_base_material,
        missing_product_codes
    )


#function to get the total booked hours and cost
def get_project_total_cost_hours(project_number):
    folder_path_po = "../Data Pool/Ecosys API Data/PO Headers"

    excel_files = [f for f in os.listdir(folder_path_po) if f.endswith(".xlsx") or f.endswith(".xls")]
    matching_files = [f for f in excel_files if str(project_number) in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found in the folder.")

    most_recent_file = get_most_recent_file(folder_path_po, matching_files)
    file_path = os.path.join(folder_path_po, most_recent_file)
    po_df = pd.read_excel(file_path)

    etreg_file_folder = "../Data Pool/Ecosys API Data/eTREG Lines"
    etreg_excel_files = [f for f in os.listdir(etreg_file_folder) if f.endswith(".xlsx") or f.endswith(".xls")]
    matching_second_files = [f for f in etreg_excel_files if str(project_number) in f]

    if not matching_second_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found in the folder."
        )

    second_most_recent_file = get_most_recent_file(etreg_file_folder, matching_second_files)
    second_file_path = os.path.join(etreg_file_folder, second_most_recent_file)
    etreg_df = pd.read_excel(second_file_path)

    sun_file_folder = "../Data Pool/Ecosys API Data/SUN Transactions"
    sun_excel_files = [f for f in os.listdir(sun_file_folder) if f.endswith(".xlsx") or f.endswith(".xls")]
    matching_third_files = [f for f in sun_excel_files if str(project_number) in f]

    if not matching_third_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' were found in the folder."
        )

    third_most_recent_file = get_most_recent_file(sun_file_folder, matching_third_files)
    third_file_path = os.path.join(sun_file_folder, third_most_recent_file)
    sun_df = pd.read_excel(third_file_path)

    total_project_cost = po_df["PO Cost"].sum()
    total_hours = etreg_df["Quantity"].sum()
    total_sun_amount = sun_df["ProjectAmount"].sum()

    return total_project_cost, total_hours, total_sun_amount



def export_complete_mto_excel():
    # TODO: Implement this function
    pass


def export_complete_mto_graphics():
    # TODO: Implement this function
    pass
