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


def yard_material_type_analyze(folder_path, project_number, material_type):
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

    required_columns = ["Pipe Base Material", "Product Code", "Total QTY to commit", "Unit Weight", "Total NET weight", "Quantity UOM", "Unit Weight UOM"]
    if all(col in df.columns for col in required_columns):
        materials = df["Pipe Base Material"].unique()

        # Group by 'Product Code' and 'Quantity UOM' and calculate the sum and average
        grouped = df.groupby(["Product Code", "Quantity UOM"]).agg(
            {"Total QTY to commit": "sum", "Unit Weight": "mean", "Total NET weight": "sum", "Pipe Base Material": "first"}
        )

        # Reset index to turn indices into columns
        grouped.reset_index(inplace=True)

        # Assuming "Unit Weight UOM" is the same for all lines, assign it directly
        grouped["Unit Weight UOM"] = "KG"

        # Define the output folder and file
        result_folder_path = "../Data Pool/DCT Process Results/Exported Result Files"
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        output_file = os.path.join(result_folder_path, f"MP{project_number} - Yard_PipingMaterial_res_{timestamp}.xlsx")

        # Write the DataFrame to an Excel file
        grouped.to_excel(output_file, index=False)
        print(f"Data has been written to {output_file}")
        ExportGraphicsbyYard.plot_material_analyze_piping(grouped, project_number)

    else:
        print("Was not possible to find the necessary fields on the file to do the calculation!")