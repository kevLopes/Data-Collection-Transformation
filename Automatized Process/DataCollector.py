# script Data Collector
import os
import pandas as pd
def DataCollectorFunc():
    # Set the search directory and keyword
    search_dir = os.path.dirname(os.path.abspath(__file__))
    keyword = "Piping"

    # Search for Excel files containing the keyword
    files = [file for file in os.listdir(search_dir) if keyword in file and file.endswith(".xlsx")]

    # Check if the Data Organize folder exists, and create it if it doesn't
    output_dir = os.path.join(search_dir, "Data Organize")
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Loop through the files and extract the Tag Number column
    for file in files:
        # Load the Excel file into a pandas dataframe
        df = pd.read_excel(os.path.join(search_dir, file))

        # Find the Tag Number column
        tag_column = None
        for column in df.columns:
            if "Tag Number" in str(column):
                tag_column = column
                break

        # If the Tag Number column was found, copy it and all rows below with data information
        if tag_column is not None:
            tag_df = df.loc[df[tag_column].notnull(), [tag_column]]

            # Save the extracted data to a new Excel file
            output_filename = os.path.join(output_dir, "Piping Organized.xlsx")
            tag_df.to_excel(output_filename, index=False)

            print(f"Extracted data from {file} and saved it to {output_filename}")
        else:
            print(f"Could not find Tag Number column in {file}")
