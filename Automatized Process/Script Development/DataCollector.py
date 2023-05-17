import os
import pandas as pd
from datetime import datetime


def data_collector_piping(project_number_v, material_type):
    print("Function to capture all Piping Data Initialized")
    global success

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Set the columns to extract
    columns_to_extract = ["Tag Number", "ID", "Project Number", "Product Code", "Commodity Code",
                          "Service Description", "Pipe Base Material", "Material", "LineNumber",
                          "SBM scope", "Total QTY to commit", "Quantity UOM", "Unit Weight",
                          "Unit Weight UOM", "Total NET weight", "SIZE"]

    # Check if the Data DW Dumber folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        print("Not possible to find folder containing Data")
        return

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        print("No excel data file for the Material " + keyword)
        return

    # Check if the Data Organize folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Piping"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        # Load the Excel file into a pandas dataframe
        df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

        # Check if the project number matches the specified one
        if df["Project Number"].astype(str).str.contains(str(project_number_v)).any():
            # Find the specified columns
            extract_columns = []
            for column in df.columns:
                if any(col in str(column) for col in columns_to_extract):
                    extract_columns.append(column)

            # If any of the specified columns were found, extract them and all rows below with data information
            if extract_columns:
                extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                # Filter rows where 'SBM scope' is equal to True
                extract_df = extract_df[extract_df['SBM scope'] == True & (extract_df['SBM scope'].notnull())]

                # Group the data by project number
                grouped_df = extract_df.groupby("Project Number")

                # Loop through the groups and save them to separate Excel files
                for project_number, group_df in grouped_df:
                    # Get the current timestamp
                    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                    # Save the group data to a new Excel file with the project number in the filename
                    output_filename = os.path.join(output_dir, f"{project_number} - Piping Organized_{timestamp}.xlsx")
                    group_df.to_excel(output_filename, index=False)

                    print(
                        f"\rExtracted data from {file} with Project Number {project_number} and saved it to {output_filename}")
                    success = True
            else:
                print(f"\rCould not find any of the specified columns in {file}")
                success = False
        else:
            print(f"No files were found for the project {project_number_v} for the {material_type} material")


# --------------------------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------------------------

def data_collector_valve(project_number_v, material_type):
    global success
    print("Function to capture all Valves Data Initialized")

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Check if the Data Pool folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        print("Data Pool folder not found on directory.")
        return

    # Set the columns to extract
    columns_to_extract = ["TAG NUMBER", "Status", "Project Number", "Product Code",
                          "Service Description", "MOC", "Bulk ID Long", "Line Number",
                          "SIZE (inch)", "General Material Description", "Quantity", "Revised", "Weight"]

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        print("No excel data file for the Material " + keyword)
        return

    # Check if the Materials Data Organized folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Valve"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        # Load the Excel file into a pandas dataframe
        df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

        # Check if the project number matches the specified one
        if df["Project Number"].astype(str).str.contains(str(project_number_v)).any():
            # Find the specified columns
            extract_columns = []
            for column in df.columns:
                if any(col in str(column) for col in columns_to_extract):
                    extract_columns.append(column)

            # If any of the specified columns were found, extract them and all rows below with data information
            if extract_columns:
                extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                # Filter rows where 'SBM scope' is equal to True
                extract_df = extract_df[extract_df['Status'] == "CONFIRMED"]

                # Group the data by project number
                grouped_df = extract_df.groupby("Project Number")

                # Loop through the groups and save them to separate Excel files
                for project_number, group_df in grouped_df:
                    # Get the current timestamp
                    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                    # Save the group data to a new Excel file with the project number in the filename
                    output_filename = os.path.join(output_dir, f"{project_number} - Valve Organized_{timestamp}.xlsx")
                    group_df.to_excel(output_filename, index=False)

                    print(
                        f"\rExtracted data from {file} with Project Number {project_number} and saved it to {output_filename}")
                    success = True
            else:
                print(f"\rCould not find any of the specified columns in {file}")
                success = False

            if not success:
                print("\rNo files were processed successfully.")
        else:
            print(f"No files were found for the project {project_number_v} for the {material_type} material")

# --------------------------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------------------------


def data_collector_bolt(project_number_v, material_type):
    print("Function to capture all Bolt Data Initialized")

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Check if the Data Pool folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        print("Data Pool folder not found.")
        return

    # Set the columns to extract
    columns_to_extract = ["Tag Number", "ID", "Project Number", "Product Code", "Commodity Code",
                          "Service Description", "Pipe Base Material", "Material", "LineNumber",
                          "SBM scope", "Qty confirmed in design", "Total QTY to commit", "Quantity UOM",
                          "Unit Weight UOM", "SIZE"]

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        print("No Excel files found for the Bolt data.")
        return

    # Check if the Materials Data Organized folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Bolt"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    success = False
    for file in files:
        # Load the Excel file into a pandas dataframe
        df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

        # Check if the project number matches the specified one
        if df["Project Number"].astype(str).str.contains(str(project_number_v)).any():
            # Find the specified columns
            extract_columns = []
            for column in df.columns:
                if any(col in str(column) for col in columns_to_extract):
                    extract_columns.append(column)

            # If any of the specified columns were found, extract them and all rows below with data information
            if extract_columns:
                extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                # Filter rows where 'SBM scope' is equal to True
                extract_df = extract_df[extract_df['SBM scope'] == True & (extract_df['SBM scope'].notnull())]

                # Group the data by project number
                grouped_df = extract_df.groupby("Project Number")

                # Loop through the groups and save them to separate Excel files
                for project_number, group_df in grouped_df:
                    # Get the current timestamp
                    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                    # Save the group data to a new Excel file with the project number in the filename
                    output_filename = os.path.join(output_dir, f"{project_number} - Bolt Organized_{timestamp}.xlsx")
                    group_df.to_excel(output_filename, index=False)

                    print(
                        f"\rExtracted data from {file} with Project Number {project_number} and saved it to {output_filename}")
                    success = True
            else:
                print(f"\rCould not find any of the specified columns in {file}")
                success = False

            if not success:
                print("\rNo files were processed successfully.")
        else:
            print(f"No files were found for the project {project_number_v} for the {material_type} material")
