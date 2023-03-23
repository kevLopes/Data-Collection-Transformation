# script Data Collector
import os
import pandas as pd
from datetime import datetime

def dataCollectorFuncPiping():

    print("Function to capture all Piping Data Initialized")

    # Set the search directory and keyword
    search_dir = os.path.dirname(os.path.abspath(__file__))
    keyword = "Piping"

    # Set the columns to extract
    columns_to_extract = ["Tag Number", "ID", "Project Number", "Product Code", "Commodity Code",
                          "Service Description", "Pipe Base Material", "Material", "LineNumber",
                          "SBM scope", "Total QTY to commit", "Quantity UOM", "Unit Weight",
                          "Unit Weight UOM", "Total NET weight", "SIZE"]

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if the Data Organize folder exists, and create it if it doesn't
    output_dir = os.path.join(search_dir, "Data Organize")
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the specified columns
    for file in files:
        # Load the Excel file into a pandas dataframe
        df = pd.read_excel(os.path.join(search_dir, file))

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

            # Save the extracted data to a new Excel file
            output_filename = os.path.join(output_dir, f"Piping Organized_{timestamp}.xlsx")
            extract_df.to_excel(output_filename, index=False)

            print(f"Extracted data from {file} and saved it to {output_filename}")
        else:
            print(f"Could not find any of the specified columns in {file}")