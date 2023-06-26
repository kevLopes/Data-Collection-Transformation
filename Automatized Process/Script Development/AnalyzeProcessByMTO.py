import os
import pandas as pd
from datetime import datetime
import ExportReportsGraphics
import ExportPDFreports


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


def analyze_by_material_type_piping(folder_path, project_number, material_type):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if str(project_number) in f and material_type in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    required_columns = ["Project Number", "Pipe Base Material", "Total QTY to commit", "Unit Weight", "Total NET weight", "Quantity UOM"]
    if all(col in df.columns for col in required_columns):
        materials = df["Pipe Base Material"].unique()

        analyze_df = pd.DataFrame(
            columns=["Project Number", "Material Type", "Quantity UOM", "Total QTY to commit", "Average Unit Weight", "Total NET weight"])

        for material in materials:
            material_data = df[df["Pipe Base Material"] == material]
            grouped_data = material_data.groupby("Quantity UOM").agg({
                "Total QTY to commit": "sum",
                "Unit Weight": "mean",
                "Total NET weight": "sum"
            }).reset_index()

            for _, row in grouped_data.iterrows():
                analyze_df = analyze_df.append({
                    "Project Number": project_number,
                    "Material Type": material,
                    "Quantity UOM": row["Quantity UOM"],
                    "Total QTY to commit": row["Total QTY to commit"],
                    "Average Unit Weight": row["Unit Weight"],
                    "Total NET weight": row["Total NET weight"]
                }, ignore_index=True)

        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Piping"
        output_file = os.path.join(result_folder_path, f"MP{project_number}_SBM_MTO_Analyze_{timestamp}.xlsx")

        # Add Quantity UOM as a separate column in the output file
        uom_mapping = df[["Quantity UOM", "Unit Weight UOM"]].drop_duplicates().set_index("Quantity UOM")["Unit Weight UOM"]
        analyze_df["Unit Weight UOM"] = analyze_df["Quantity UOM"].map(uom_mapping)

        analyze_df.to_excel(output_file, index=False)
        ExportReportsGraphics.sbm_scope_mto_plot_piping_analyze(analyze_df, project_number)
        print("SBM Scope MTO Data analysis has been saved.")
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")


#Read Data of Piping SBM Scope for PDF report
def sbm_scope_piping_report(folder_path, project_number, material_type):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if str(project_number) in f and material_type in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    required_columns = ["Project Number", "Pipe Base Material", "Total QTY to commit", "Unit Weight", "Total NET weight", "Quantity UOM"]

    if all(col in df.columns for col in required_columns):

        # Filter the data based on the "SBM Scope" column
        df = df[df['SBM scope'] == True]

        # Get the project number
        project_number = df['Project Number'].iloc[0]
        ExportPDFreports.generate_pdf_piping_sbm_scope(df, project_number, "Bulk Team MTO Data")
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")


#Read Data of Piping YARD Scope for PDF report
def yard_scope_piping_report(folder_path, project_number, material_type):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if str(project_number) in f and material_type in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found.")

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    required_columns = ["Project Number", "Pipe Base Material", "Total QTY to commit", "Unit Weight", "Total NET weight", "Quantity UOM"]

    if all(col in df.columns for col in required_columns):

        # Filter the data based on the "SBM Scope" column
        df = df[df['SBM scope'] == False]

        # Get the project number
        #project_number = df['Project Number'].iloc[0]
        ExportPDFreports.generate_pdf_piping_yard_scope(df, project_number, "Bulk Team MTO Data")
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")