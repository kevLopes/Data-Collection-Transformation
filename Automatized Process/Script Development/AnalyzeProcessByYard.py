import os
import pandas as pd
from datetime import datetime
import ExportGraphicsbyYard


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


def yard_piping_material_type_analyze(folder_path, project_number, material_type):
    excel_files = [
        f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    matching_files = [
        f for f in excel_files if str(project_number) in f and material_type in f and "Yard" in f
    ]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found."
        )

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    required_columns = ["Pipe Base Material", "Product Code", "Total QTY to commit", "Unit Weight", "Total NET weight", "Quantity UOM", "Unit Weight UOM"]

    if all(col in df.columns for col in required_columns):

        # Group by 'Product Code' and 'Quantity UOM' and calculate the sum and average
        grouped = df.groupby(["Product Code", "Quantity UOM"]).agg(
            {"Total QTY to commit": "sum", "Unit Weight": "mean", "Total NET weight": "sum", "Pipe Base Material": "first"}
        )

        # Reset index to turn indices into columns
        grouped.reset_index(inplace=True)

        # Assuming "Unit Weight UOM" is the same for all lines, assign it directly
        grouped["Unit Weight UOM"] = "KG"

        # Define the output folder and file
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Yard"
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        output_file = os.path.join(result_folder_path, f"MP{project_number} - Yard PipingMaterial_{timestamp}.xlsx")

        # Write the DataFrame to an Excel file
        grouped.to_excel(output_file, index=False)
        print(f"Data has been written to {output_file}")

        ExportGraphicsbyYard.plot_material_analyze_piping(grouped, project_number)

    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")


#                       ------------------------ Valve --------------------


def yard_valve_material_type_analyze(folder_path, project_number, material_type):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if str(project_number) in f and material_type in f and "Yard" in f]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found."
        )

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    required_columns = ["General Material Description", "Product Code", "Quantity", "Weight", "SIZE (inch)", "Status", "Remarks"]
    if all(col in df.columns for col in required_columns):
        # Group by 'Product Code' and calculate the sum and average
        grouped = df.groupby(["Product Code"]).agg(
            {"General Material Description": "first", "Quantity": "sum", "Weight": "sum", "SIZE (inch)": "mean"}
        )

        # Reset index to turn indices into columns
        grouped.reset_index(inplace=True)

        # Define the output folder and file
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Yard"
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        output_file = os.path.join(result_folder_path, f"MP{project_number} - Yard ValveMaterial_{timestamp}.xlsx")

        # Write the DataFrame to an Excel file
        grouped.to_excel(output_file, index=False)
        print(f"Data has been written to {output_file}")

        ExportGraphicsbyYard.plot_material_analyze_valves(grouped, project_number)
    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")

#                       ------------------------ Bolt --------------------


def yard_bolt_material_type_analyze(folder_path, project_number, material_type):
    excel_files = [
        f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    matching_files = [
        f for f in excel_files if str(project_number) in f and material_type in f and "Yard" in f
    ]

    if not matching_files:
        raise FileNotFoundError(
            f"No files containing the project number '{project_number}' and material type '{material_type}' were found."
        )

    most_recent_file = get_most_recent_file(folder_path, matching_files)
    file_path = os.path.join(folder_path, most_recent_file)
    df = pd.read_excel(file_path)

    required_columns = ["Pipe Base Material", "Product Code", "Total QTY to commit", "Qty confirmed in design", "SIZE", "SBM scope", "Tag Number"]
    if all(col in df.columns for col in required_columns):
        # Calculate total quantity to commit and quantity confirmed in design
        df["Total Quantity"] = df["Total QTY to commit"].astype(float)
        df["Quantity Confirmed"] = df["Qty confirmed in design"].astype(float)

        # Group by Pipe Base Material and Product Code
        grouped = df.groupby(["Pipe Base Material", "Product Code"]).agg({
            "Total Quantity": "sum",
            "Quantity Confirmed": "sum"
        }).reset_index()

        # Create Quantity Difference column
        grouped["Quantity Difference"] = grouped["Total Quantity"] - grouped["Quantity Confirmed"]

        # Add Quantity UOM column
        grouped["Quantity UOM"] = "PCE"

        # Rename Pipe Base Material column to General Material Description
        grouped = grouped.rename(columns={"Pipe Base Material": "General Material Description"})

        # Define the output folder and file
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files/Yard"
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        output_file = os.path.join(result_folder_path, f"MP{project_number} - Yard BoltMaterial_{timestamp}.xlsx")

        # Write the DataFrame to an Excel file
        grouped.to_excel(output_file, index=False)
        #print(f"Data has been written to {output_file}")

        print(f"Bolt material analysis for project {project_number} with material type {material_type} has been exported successfully.")

        ExportGraphicsbyYard.plot_material_analyze_bolts(grouped, project_number)

    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")
