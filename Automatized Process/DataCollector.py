# script Data Collector
import os
import pandas as pd
from datetime import datetime

def dataCollectorFuncPiping():

    print("Function to capture all Piping Data Initialized")

    global success

    # Set the search directory and keyword
    search_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(search_dir, "Data Pool")
    keyword = "Piping"

    # Set the columns to extract
    columns_to_extract = ["Tag Number", "ID", "Project Number", "Product Code", "Commodity Code",
                          "Service Description", "Pipe Base Material", "Material", "LineNumber",
                          "SBM scope", "Total QTY to commit", "Quantity UOM", "Unit Weight",
                          "Unit Weight UOM", "Total NET weight", "SIZE"]

    # Check if the Data DW Dumber folder exists, and display an error message if it doesn't
    if not os.path.exists(data_dir):
        print("Not possible to find folder containing Data")
        return

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        print("No excel data file for the Material " + keyword)
        return

    # Check if the Data Organize folder exists, and create it if it doesn't
    output_dir = os.path.join(search_dir, "Materials Data Organized")
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
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

            # Get the project number from the dataframe
            project_number = extract_df["Project Number"].iloc[0]

            # Get the current timestamp
            timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

            # Save the extracted data to a new Excel file with the project number in the filename
            output_filename = os.path.join(output_dir, f"{project_number} - Piping Organized_{timestamp}.xlsx")
            extract_df.to_excel(output_filename, index=False)

            print(f"Extracted data from {file} and saved it to {output_filename}")
            success = True
        else:
            print(f"Could not find any of the specified columns in {file}")
            success = False

        if not success:
            print("No files were processed successfully.")

#TO BE ADJUSTED AND TESTED - VALVES
def dataCollectorFuncValve():

    global success
    print("Function to capture all Valves Data Initialized")

    # Set the search directory and keyword
    search_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(search_dir, "Data Pool")
    keyword = "Valves"

    # Set the columns to extract
    columns_to_extract = ["Tag Number", "ID", "Project Number", "Product Code", "Commodity Code",
                          "Service Description", "Pipe Base Material", "Material", "LineNumber",
                          "SBM scope", "Total QTY to commit", "Quantity UOM", "Unit Weight",
                          "Unit Weight UOM", "Total NET weight", "SIZE"]

    # Check if the Data DW Dumber folder exists, and display an error message if it doesn't
    if not os.path.exists(data_dir):
        print("Not possible to find folder containing Data")
        return

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        print("No excel data file for the Material " + keyword)
        return

    # Check if the Data Organize folder exists, and create it if it doesn't
    output_dir = os.path.join(search_dir, "Materials Data Organized")
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
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

            # Get the project number from the dataframe
            project_number = extract_df["Project Number"].iloc[0]

            # Get the current timestamp
            timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

            # Save the extracted data to a new Excel file with the project number in the filename
            output_filename = os.path.join(output_dir, f"{project_number} - Valves Organized_{timestamp}.xlsx")
            extract_df.to_excel(output_filename, index=False)

            print(f"Extracted data from {file} and saved it to {output_filename}")
            success = True
        else:
            print(f"Could not find any of the specified columns in {file}")
            success = False

    if not success:
        print("No files were processed successfully.")


#TO BE ADJUSTED AND TESTED - BOLT
def dataCollectorFuncBolt():

    print("Function to capture all Bolt Data Initialized")

    global success

    # Set the search directory and keyword
    search_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(search_dir, "Data Pool")
    keyword = "Bolt"

    # Check if the Data DW Dumber folder exists, and display an error message if it doesn't
    if not os.path.exists(data_dir):
        print("Not possible to find folder containing Data")
        return

    # Set the columns to extract
    columns_to_extract = ["Tag Number", "ID", "Project Number", "Product Code", "Commodity Code",
                          "Service Description", "Pipe Base Material", "Material", "LineNumber",
                          "SBM scope", "Total QTY to commit", "Quantity UOM", "Unit Weight",
                          "Unit Weight UOM", "Total NET weight", "SIZE"]

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if any Excel files were found with the specified keyword
    if not files:
        print("No excel data file for the Material " + keyword)
        return

    # Check if the Data Organize folder exists, and create it if it doesn't
    output_dir = os.path.join(search_dir, "Materials Data Organized")
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    success = False
    for file in files:
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

            # Get the project number from the dataframe
            project_number = extract_df["Project Number"].iloc[0]

            # Get the current timestamp
            timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

            # Save the extracted data to a new Excel file with the project number in the filename
            output_filename = os.path.join(output_dir, f"{project_number} - Bolt Organized_{timestamp}.xlsx")
            extract_df.to_excel(output_filename, index=False)

            print(f"Extracted data from {file} and saved it to {output_filename}")
            success = True
        else:
            print(f"Could not find any of the specified columns in {file}")
            success = False

    if not success:
        print("No files were processed successfully.")

