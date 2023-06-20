import os
import pandas as pd
from datetime import datetime
import logging
logging.basicConfig(level=logging.INFO)


#Data Collector and Transformation SBM Scope Piping
def data_collector_piping(project_number, material_type):
    logging.info("Function to capture all Piping Data Initialized")

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
        logging.error("Not possible to find folder containing Data")
        return False

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        logging.error("No excel data file for the Material " + keyword)
        return False

    # Check if the Data Organize folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Piping"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        try:
            # Load the Excel file into a pandas dataframe
            df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

            # Check if the project number matches the specified one
            if df["Project Number"].astype(str).str.contains(str(project_number)).any():
                # Find the specified columns
                extract_columns = []
                for column in df.columns:
                    if any(col in str(column) for col in columns_to_extract):
                        extract_columns.append(column)

                # If any of the specified columns were found, extract them and all rows below with data information
                if extract_columns:
                    extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                    # Filter rows where 'SBM scope' is equal to True
                    extract_df = extract_df[(extract_df['SBM scope'] == True) & (extract_df['SBM scope'].notnull())]

                    # Group the data by project number
                    grouped_df = extract_df.groupby("Project Number")

                    # Loop through the groups and save them to separate Excel files
                    for project_number, group_df in grouped_df:
                        # Get the current timestamp
                        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                        # Save the group data to a new Excel file with the project number in the filename
                        output_filename = os.path.join(output_dir, f"{project_number} - Piping Organized_{timestamp}.xlsx")
                        group_df.to_excel(output_filename, index=False)

                        logging.info(
                            f"Extracted data from {file} with Project Number {project_number} and saved it to {output_filename}")
                        return True
                else:
                    logging.error(f"Could not find any of the specified columns in {file}")
                    return False
            else:
                logging.error(f"No files were found for the project {project_number} for the {material_type} material")
                return False

        except Exception as e:
            logging.error(f"Failed to process {file} with error {e}")
            return False


#Data Collector and Transformation YARD Scope Piping
def data_collector_piping_yard(project_number, material_type):

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
        logging.error("Not possible to find folder containing Data")
        return False

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        logging.error("No excel data file for the Material " + keyword)
        return False

    # Check if the Data Organize folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Piping"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        try:
            # Load the Excel file into a pandas dataframe
            df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

            # Check if the project number matches the specified one
            if df["Project Number"].astype(str).str.contains(str(project_number)).any():
                # Find the specified columns
                extract_columns = []
                for column in df.columns:
                    if any(col in str(column) for col in columns_to_extract):
                        extract_columns.append(column)

                # If any of the specified columns were found, extract them and all rows below with data information
                if extract_columns:
                    extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                    # Filter rows where 'YARD scope' is equal to True
                    extract_df = extract_df[(extract_df['SBM scope'] == False) & (extract_df['SBM scope'].notnull())]

                    # Group the data by project number
                    grouped_df = extract_df.groupby("Project Number")

                    # Loop through the groups and save them to separate Excel files
                    for project_number, group_df in grouped_df:
                        # Get the current timestamp
                        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                        # Save the group data to a new Excel file with the project number in the filename
                        output_filename = os.path.join(output_dir, f"{project_number} - Yard Piping Organized_{timestamp}.xlsx")
                        group_df.to_excel(output_filename, index=False)

                        logging.info(
                            f"Extracted data from {file} with Project Number {project_number} and saved it to {output_filename} (YARD Scope)")
                        return True
                else:
                    logging.error(f"Could not find any of the specified columns in {file} (YARD Scope)")
                    return False
            else:
                logging.error(f"No files were found for the project {project_number} for the {material_type} material (YARD Scope)")
                return False

        except Exception as e:
            logging.error(f"Failed to process {file} with error (YARD Scope) {e}")
            return False


# --------------------------------------------------------------------------------------------------------------------
# --------------------------------------VALVE--------------------------------------------------------------------


#Data Collector and Transformation SBM Scope Valve
def data_collector_valve(project_number, material_type):
    logging.info("Function to capture all Valves Data Initialized")

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Set the columns to extract
    columns_to_extract = ["TAG NUMBER", "Status", "Project Number", "Product Code",
                          "Service Description", "MOC", "Bulk ID Long", "Line Number",
                          "SIZE (inch)", "General Material Description", "Quantity", "Revised", "Weight", "Remarks"]

    # Check if the Data Pool folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        logging.error("Data Pool folder not found on directory.")
        return False

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        logging.error("No excel data file for the Material " + keyword)
        return False

    # Check if the Materials Data Organized folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Valve"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        try:
            # Load the Excel file into a pandas dataframe
            df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

            # Check if the project number matches the specified one
            if df["Project Number"].astype(str).str.contains(str(project_number)).any():
                # Find the specified columns
                extract_columns = []
                for column in df.columns:
                    if any(col in str(column) for col in columns_to_extract):
                        extract_columns.append(column)

                # If any of the specified columns were found, extract them and all rows below with data information
                if extract_columns:
                    extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                    # Filter rows where 'Status' is equal to "CONFIRMED" and 'Remarks' does not contain "Yard"
                    extract_df = extract_df[(extract_df['Status'] == "CONFIRMED") & (~extract_df['Remarks'].str.contains("Yard", na=False))]

                    # Group the data by project number
                    grouped_df = extract_df.groupby("Project Number")

                    # Loop through the groups and save them to separate Excel files
                    for project_number, group_df in grouped_df:
                        # Get the current timestamp
                        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                        # Save the group data to a new Excel file with the project number in the filename
                        output_filename = os.path.join(output_dir, f"{project_number} - Valve Organized_{timestamp}.xlsx")
                        group_df.to_excel(output_filename, index=False)

                        logging.info(
                            f"Extracted data from {file} with Project Number {project_number} and saved it to {output_filename}")
                        return True
                else:
                    logging.error(f"Could not find any of the specified columns in {file}")
                    return False
            else:
                logging.error(f"No files were found for the project {project_number} for the {material_type} material")
                return False

        except Exception as e:
            logging.error(f"Failed to process {file} with error {e}")
            return False


#Data Collector and Transformation YARD Scope Valve
def data_collector_valve_yard(project_number, material_type):

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Set the columns to extract
    columns_to_extract = ["TAG NUMBER", "Status", "Project Number", "Product Code",
                          "Service Description", "MOC", "Bulk ID Long", "Line Number",
                          "SIZE (inch)", "General Material Description", "Quantity", "Revised", "Weight", "Remarks"]

    # Check if the Data Pool folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        logging.error("Data Pool folder not found on directory.")
        return False

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        logging.error("No excel data file for the Material " + keyword)
        return False

    # Check if the Materials Data Organized folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Valve"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        try:
            # Load the Excel file into a pandas dataframe
            df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

            # Check if the project number matches the specified one
            if df["Project Number"].astype(str).str.contains(str(project_number)).any():
                # Find the specified columns
                extract_columns = []
                for column in df.columns:
                    if any(col in str(column) for col in columns_to_extract):
                        extract_columns.append(column)

                # If any of the specified columns were found, extract them and all rows below with data information
                if extract_columns:
                    extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                    # Filter rows where 'Status' is equal to "OUT OF SCOPE" or 'Remarks' contains "Yard"
                    extract_df = extract_df[(extract_df['Status'] == "OUT OF SCOPE") | (extract_df['Remarks'].str.contains("Yard", na=False))]

                    # Group the data by project number
                    grouped_df = extract_df.groupby("Project Number")

                    # Loop through the groups and save them to separate Excel files
                    for project_number, group_df in grouped_df:
                        # Get the current timestamp
                        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                        # Save the group data to a new Excel file with the project number in the filename
                        output_filename = os.path.join(output_dir, f"{project_number} - Yard Valve Organized_{timestamp}.xlsx")
                        group_df.to_excel(output_filename, index=False)

                        logging.info(
                            f"Extracted data from {file} with Project Number {project_number} and saved it to {output_filename} (YARD Scope)")
                        return True
                else:
                    logging.error(f"Could not find any of the specified columns in {file} (YARD Scope)")
                    return False
            else:
                logging.error(f"No files were found for the project {project_number} for the {material_type} material (YARD Scope)")
                return False

        except Exception as e:
            logging.error(f"Failed to process {file} with error (YARD Scope) {e}")
            return False


# --------------------------------------------------------------------------------------------------------------------
# -------------------------------------BOLT-----------------------------------------------------------------------


#Data Collector and Transformation YARD Scope Bolt
def data_collector_bolt_yard(project_number, material_type):

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Set the columns to extract
    columns_to_extract = ["Tag Number", "ID", "Project Number", "Product Code", "Commodity Code",
                          "Service Description", "Pipe Base Material", "Material", "LineNumber",
                          "SBM scope", "Qty confirmed in design", "Total QTY to commit", "Quantity UOM",
                          "SIZE"]

    # Check if the Data Pool folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        logging.error("Data Pool folder not found.")
        return False

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        logging.error("No Excel files found for the Bolt data.")
        return False

    # Check if the Materials Data Organized folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Bolt"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        try:
            # Load the Excel file into a pandas dataframe
            df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

            # Check if the project number matches the specified one
            if df["Project Number"].astype(str).str.contains(str(project_number)).any():
                # Find the specified columns
                extract_columns = []
                for column in df.columns:
                    if any(col in str(column) for col in columns_to_extract):
                        extract_columns.append(column)

                # If any of the specified columns were found, extract them and all rows below with data information
                if extract_columns:
                    extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                    # Filter rows where 'YARD scope' is equal to True
                    extract_df = extract_df[extract_df['SBM scope'] == False & (extract_df['SBM scope'].notnull())]

                    # Group the data by project number
                    grouped_df = extract_df.groupby("Project Number")

                    # Loop through the groups and save them to separate Excel files
                    for project_number, group_df in grouped_df:
                        # Get the current timestamp
                        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                        # Save the group data to a new Excel file with the project number in the filename
                        output_filename = os.path.join(output_dir, f"{project_number} - Yard Bolt Organized_{timestamp}.xlsx")
                        group_df.to_excel(output_filename, index=False)

                        logging.info(
                            f"Extracted data from {file} with Project Number {project_number} and saved it to {output_filename} (YARD Scope)")
                        return True
                else:
                    logging.error(f"Could not find any of the specified columns in {file} (YARD Scope)")
                    return False
            else:
                logging.error(f"No files were found for the project {project_number} for the {material_type} material (YARD Scope)")
                return False

        except Exception as e:
            logging.error(f"Failed to process file {file} due to error (YARD Scope): {e}")
            return False

    logging.error("No files were processed successfully (YARD Scope).")
    return False


#Data Collector and Transformation SBM Scope Bolt
def data_collector_bolt(project_number, material_type):
    logging.info("Function to capture all Bolt Data Initialized")

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Set the columns to extract
    columns_to_extract = ["Tag Number", "ID", "Project Number", "Product Code", "Commodity Code",
                          "Service Description", "Pipe Base Material", "Material", "LineNumber",
                          "SBM scope", "Qty confirmed in design", "Total QTY to commit", "Quantity UOM",
                          "SIZE"]

    # Check if the Data Pool folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        logging.error("Data Pool folder not found.")
        return False

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        logging.error("No Excel files found for the Bolt data.")
        return False

    # Check if the Materials Data Organized folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Bolt"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        try:
            # Load the Excel file into a pandas dataframe
            df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

            # Check if the project number matches the specified one
            if df["Project Number"].astype(str).str.contains(str(project_number)).any():
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

                        logging.info(
                            f"Extracted data from {file} with Project Number {project_number} and saved it to {output_filename}")
                        return True
                else:
                    logging.error(f"Could not find any of the specified columns in {file}")
                    return False
            else:
                logging.error(f"No files were found for the project {project_number} for the {material_type} material")
                return False

        except Exception as e:
            logging.error(f"Failed to process file {file} due to error: {e}")
            return False

    logging.error("No files were processed successfully.")
    return False

# --------------------------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------------------------


def data_collector_structure(project_number, material_type):
    logging.info("Function to capture all Structure Data Initialized")

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Set the columns to extract
    columns_to_extract = ["Tag Number", "Project", "Product Code", "Commodity Code",
                          "Service Description", "Thickness", "Material", "Wastage Quantity",
                          "Required Qty", "Unit Weight", "Total QTY to commit", "Quantity UOM",
                          "Unit Weight UOM", "Total NET weight", "Quantity Including Wastage", "Total Gross Weight", "Total Gross Weight UOM"]

    # Check if the Data Pool folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        logging.error("Data Pool folder not found.")
        return False

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        logging.error("No Excel files found for the Structure data.")
        return False

    # Check if the Materials Data Organized folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Structure"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        try:
            # Load the Excel file into a pandas dataframe
            df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

            # Check if the project number matches the specified one
            if df["Project"].astype(str).str.contains(str(project_number)).any():
                # Find the specified columns
                extract_columns = []
                for column in df.columns:
                    if any(col in str(column) for col in columns_to_extract):
                        extract_columns.append(column)

                # If any of the specified columns were found, extract them and all rows below with data information
                if extract_columns:
                    extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                    # Group the data by project number
                    grouped_df = extract_df.groupby("Project")

                    # Loop through the groups and save them to separate Excel files
                    for project_number, group_df in grouped_df:
                        # Get the current timestamp
                        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                        # Save the group data to a new Excel file with the project number in the filename
                        output_filename = os.path.join(output_dir, f"{project_number} - Structure Organized_{timestamp}.xlsx")
                        group_df.to_excel(output_filename, index=False)

                        logging.info(
                            f"Extracted data from {file} with Project Number {project_number} and saved it to {output_filename}")
                        return True
                else:
                    logging.error(f"Could not find any of the specified columns in {file}")
                    return False
            else:
                logging.error(f"No files were found for the project {project_number} for the {material_type} material")
                return False
        except Exception as e:
            logging.error(f"Failed to process file {file} due to error: {e}")
            return False

    logging.error("No files were processed successfully.")
    return False

# --------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------BEND----------------------------------------------------------------


def data_collector_bend(project_number, material_type):
    logging.info("Function to capture all Bend Data Initialized")

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Set the columns to extract
    columns_to_extract = ["Tag no.", "Size NPS", "Description", "Material - MDS", "MDS", "Module", "Design Press", "Status"]

    # Check if the Data Pool folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        logging.error("Data Pool folder not found.")
        return False

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        logging.error("No Excel files found for the Bend data.")
        return False

    # Check if the Materials Data Organized folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Bend"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        try:
            # Load the Excel file into a pandas dataframe
            df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

            # Find the specified columns
            extract_columns = []
            for column in df.columns:
                if any(col in str(column) for col in columns_to_extract):
                    extract_columns.append(column)

            # If any of the specified columns were found, extract them and all rows below with data information
            if extract_columns:
                extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                # Get the current timestamp
                timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                # Save the data to a new Excel file with the project number in the filename
                output_filename = os.path.join(output_dir, f"SO{project_number} - Bend Organized_{timestamp}.xlsx")
                extract_df.to_excel(output_filename, index=False)

                logging.info(
                    f"Extracted data from {file} with Project Number {project_number} and saved it to {output_filename}")
                return True
            else:
                logging.error(f"Could not find any of the specified columns in {file}")
                return False
        except Exception as e:
            logging.error(f"Failed to process file {file} due to error: {e}")
            return False

    logging.error("No files were processed successfully.")
    return False

# --------------------------------------------------------------------------------------------------------------------
# -------------------------------------------SPECIAL PIPING-----------------------------------------------------------


#Data Collector and Transformation SBM Scope Special Piping
def data_collector_specialpip(project_number, material_type):
    logging.info("Function to capture all Special Piping Data Initialized")

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Set the columns to extract
    columns_to_extract = ["TagNumber", "Size (Inch)", "Description", "Weight", "Service", "Remarks", "PO Number", "Qty", "Linenumber"]

    # Check if the Data Pool folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        logging.error("Data Pool folder not found.")
        return False

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        logging.error("No Excel files found for the Special Piping data.")
        return False

    # Check if the Materials Data Organized folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Special Piping"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        try:
            # Load the Excel file into a pandas dataframe
            df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

            # Filter out the rows where 'PO Number' is 'BY YARD'
            df = df[df['PO Number'] != 'BY YARD']

            # Find the specified columns
            extract_columns = []
            for column in df.columns:
                if any(col in str(column) for col in columns_to_extract):
                    extract_columns.append(column)

            # If any of the specified columns were found, extract them and all rows below with data information
            if extract_columns:
                extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                # Get the current timestamp
                timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                # Save the data to a new Excel file with the project number in the filename
                output_filename = os.path.join(output_dir, f"SO{project_number} - SPC Piping Organized_{timestamp}.xlsx")
                extract_df.to_excel(output_filename, index=False)

                logging.info(
                    f"Extracted data from {file} with Project Number {project_number} and saved it to {output_filename}")
                return True
            else:
                logging.error(f"Could not find any of the specified columns in {file}")
                return False
        except Exception as e:
            logging.error(f"Failed to process file {file} due to error: {e}")
            return False

    logging.error("No files were processed successfully.")
    return False


#Data Collector and Transformation YARD Scope Special Piping
def data_collector_specialpip_yard(project_number, material_type):

    # Set the search directory and keyword
    search_dir = "../Data Pool/Data Hub Materials"
    keyword = material_type

    # Set the columns to extract
    columns_to_extract = ["TagNumber", "Size (Inch)", "Description", "Weight", "Service", "Remarks", "PO Number", "Qty", "Linenumber"]

    # Check if the Data Pool folder exists, and display an error message if it doesn't
    if not os.path.exists(search_dir):
        logging.error("Data Pool folder not found.")
        return False

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        logging.error("No Excel files found for the Special Piping data.")
        return False

    # Check if the Materials Data Organized folder exists, and create it if it doesn't
    output_dir = "../Data Pool/Material Data Organized/Special Piping"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        try:
            # Load the Excel file into a pandas dataframe
            df = pd.read_excel(os.path.join(search_dir, file), engine='openpyxl')

            # Filter the rows where 'PO Number' is 'BY YARD'
            df = df[df['PO Number'] == 'BY YARD']

            # Find the specified columns
            extract_columns = []
            for column in df.columns:
                if any(col in str(column) for col in columns_to_extract):
                    extract_columns.append(column)

            # If any of the specified columns were found, extract them and all rows below with data information
            if extract_columns:
                extract_df = df.loc[df[extract_columns].notnull().any(axis=1), extract_columns]

                # Get the current timestamp
                timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

                # Save the data to a new Excel file with the project number in the filename
                output_filename = os.path.join(output_dir, f"SO{project_number} - Yard SPC_Pip Organized_{timestamp}.xlsx")
                extract_df.to_excel(output_filename, index=False)

                logging.info(
                    f"Extracted data from {file} with Project Number {project_number} and saved it to {output_filename} (YARD Scope)")
                return True
            else:
                logging.error(f"Could not find any of the specified columns in {file} (YARD Scope)")
                return False
        except Exception as e:
            logging.error(f"Failed to process file {file} due to error (YARD Scope): {e}")
            return False

    logging.error("No files were processed successfully (YARD Scope).")
    return False
