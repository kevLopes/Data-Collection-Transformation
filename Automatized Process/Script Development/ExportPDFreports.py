import os
from datetime import date
from fpdf import FPDF
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns


# -------------------------- Piping -----------------------------
# SBM Scope
def generate_pdf_piping_sbm_scope(data, project_number, file_name):
    # Convert Total NET weight to TON
    data['Total NET weight'] = data['Total NET weight'] / 1000

    # Group the data by material type
    grouped_data = data.groupby('Pipe Base Material')

    # Initialize PDF object
    pdf = FPDF()

    # Set up the PDF
    pdf.set_title("MTO Piping Analyze - SBM Scope")
    pdf.add_page()

    # Set font and font size for the title
    pdf.set_font("Arial", "B", 16)

    # Add logo
    logo_path = "../Data Pool/DCT Process Results/PDF Reports/Content/SBM.png"
    pdf.image(logo_path, x=10, y=10, w=30)

    # Add title
    pdf.ln(20)
    pdf.cell(0, 10, "MP17033 - Unity MTO Piping Analyze - SBM Scope", ln=True, align="C")

    # Add date of analysis
    pdf.ln(3)
    pdf.set_font("Arial", "I", 8)
    pdf.cell(0, 10, f"Date of Analysis: {date.today().strftime('%d-%b-%Y')}", ln=True, align="R")

    pdf.ln(5)

    # Add summary statistics
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total Number of Records: {len(data)}", ln=True, align="L")
    pdf.cell(0, 10, f"Overall Total NET weight (TON): {data['Total NET weight'].sum()}", ln=True, align="L")
    pdf.cell(0, 10, f"Average Unit Weight (Kg): {data['Unit Weight'].mean()}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity Committed: {data['Total QTY to commit'].sum()}", ln=True, align="L")

    pdf.ln(10)  # Add a new line

    # Add material type breakdown (Page 1)
    pdf.set_font("Arial", "B", 13)
    pdf.cell(0, 10, "Material Type Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Create a table for material type breakdown
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)  # Set background color for header row

    # Header row
    pdf.cell(70, 10, "Material Type", border=1, fill=True, align="C")
    pdf.cell(30, 10, "Count", border=1, fill=True, align="C")
    pdf.ln(10)  # Add a new line

    pdf.set_font("Arial", "", 9)
    material_types = []
    material_counts = []
    for material_type, group in grouped_data:
        # Calculate the required height for the cell based on the length of the material type name
        material_type_height = 10 + (pdf.get_string_width(material_type) // 55) * 10

        # Adjust the row height if the material type name exceeds the column width
        row_height = max(material_type_height, 10)  # Minimum row height is 10

        # Draw the material type cell with adjusted height
        pdf.cell(70, row_height, material_type, border=1, align="L")
        pdf.cell(30, row_height, str(len(group)), border=1, align="C")
        pdf.ln(row_height)  # Add a new line

        material_types.append(material_type)
        material_counts.append(len(group))

    pdf.add_page()  # Add a new page for Material Type Breakdown (Pie Chart) and Data Source and Filters

    # Add material type breakdown (Page 2)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Material Type Breakdown (Pie Chart):", ln=True, align="L")
    pdf.ln(5)  # Add space

    plt.figure(figsize=(13, 13))
    plt.pie(material_counts, labels=material_types, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the pie chart as an image in the Content folder
    chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/material_type_breakdown.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()

    # Add the image to the PDF report
    pdf.image(chart_path, x=50, y=pdf.get_y(), w=100)

    pdf.ln(100)  # Add space

    # Add a new pie chart for Total NET weight by material type
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Total NET Weight Breakdown by Material Type:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Calculate the total NET weight for each material type
    total_net_weights = [group['Total NET weight'].sum() for material_type, group in grouped_data]

    plt.figure(figsize=(17, 17))
    plt.pie(total_net_weights, labels=material_types, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the new pie chart as an image in the Content folder
    net_weight_chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/total_net_weight_breakdown.png"
    plt.savefig(net_weight_chart_path, bbox_inches='tight')
    plt.close()

    # Add the new image to the PDF report
    pdf.image(net_weight_chart_path, x=50, y=pdf.get_y(), w=100)

    # Add data source and filters information
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Source and Filters:", ln=True, align="L")
    pdf.ln(5)  # Add space

    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Data Source: {file_name}", ln=True, align="L")
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Filters Applied:", ln=True, align="L")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, "- SBM scope = True", ln=True, align="L")

    pdf.ln(10)  # Add space

    # Add data breakdown and insights
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    total_qty_pce = 0
    total_qty_mtrs = 0

    # Add breakdown for each material type
    pdf.set_font("Arial", "B", 11)
    for material_type, group in grouped_data:
        pdf.set_font("Arial", "BI", 11)
        pdf.cell(0, 10, f"Material Type: {material_type}", ln=True, align="L")
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 10, f"Total NET weight (TON): {group['Total NET weight'].sum()}", ln=True, align="L")
        pdf.cell(0, 10, f"Average Unit Weight (Kg): {group['Unit Weight'].mean()}", ln=True, align="L")

        # Filter data by Quantity UOM METER and PCS
        filtered_data_meter = group[group['Quantity UOM'] == 'METER']
        filtered_data_pcs = group[group['Quantity UOM'] == 'PCE']

        # If data exists for METER and PCS, then print the corresponding total quantities
        if not filtered_data_meter.empty:
            pdf.cell(0, 10, f"Total QTY to commit (METER): {filtered_data_meter['Total QTY to commit'].sum()}", ln=True,
                     align="L")
            total_qty_mtrs += filtered_data_meter['Total QTY to commit'].sum()

        if not filtered_data_pcs.empty:
            pdf.cell(0, 10, f"Total QTY to commit (PCE): {filtered_data_pcs['Total QTY to commit'].sum()}", ln=True,
                     align="L")
            total_qty_pce += filtered_data_pcs['Total QTY to commit'].sum()

        pdf.ln(5)  # Add a new line

    pdf.add_page()

    # Add insights
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 10, "Quantities Checks:", ln=True, align="L")
    pdf.ln(3)  # Add space
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10, f"Total Quantity in Meters: {total_qty_mtrs}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity in Pieces: {total_qty_pce}", ln=True, align="L")
    pdf.ln(5)  # Add space
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 10, "Insights:", ln=True, align="L")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10,
             "Carbon steel is the most prevalent material type, accounting for a significant portion of the total records.",
             ln=True, align="L")
    pdf.cell(0, 10, "Chrome-moly and nickel alloy materials have relatively low representation in the dataset.",
             ln=True, align="L")
    pdf.cell(0, 10,
             "The total net weight varies across different materials, with carbon steel and duplex stainless steel having the highest values.",
             ln=True, align="L")
    pdf.cell(0, 10,
             "The average unit weight also shows variation, with chrome-moly and non-metallic gaskets having higher average weights.",
             ln=True, align="L")
    pdf.cell(0, 10,
             "The quantity committed differs based on the material type and is reported in both meters and pieces.",
             ln=True, align="L")

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Piping"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"{project_number}_MTOPipingReport_SBMscope.pdf")
    pdf.output(pdf_path)


# SBM Scope
def generate_pdf_piping_sbm_scope1(data, project_number, file_name):
    # Convert Total NET weight to TON
    data['Total NET weight'] = data['Total NET weight'] / 1000

    # Group the data by material type
    grouped_data = data.groupby('Pipe Base Material')

    # Initialize PDF object
    pdf = FPDF()

    # Set up the PDF
    pdf.set_title("MTO Piping Analyze - SBM Scope")
    pdf.add_page()

    # Set font and font size for the title
    pdf.set_font("Arial", "B", 16)

    # Add title
    pdf.cell(0, 10, "Piping MTO Analyze - SBM Scope", ln=True, align="C")

    # Add date of analysis
    pdf.set_font("Arial", "I", 12)
    pdf.cell(0, 10, f"Date of Analysis: {date.today().strftime('%d-%b-%Y')}", ln=True, align="R")

    # Add summary statistics
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total Number of Records: {len(data)}", ln=True, align="L")
    pdf.cell(0, 10, f"Overall Total NET weight (TON): {data['Total NET weight'].sum()}", ln=True, align="L")
    pdf.cell(0, 10, f"Average Unit Weight (Kg): {data['Unit Weight'].mean()}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity Committed: {data['Total QTY to commit'].sum()}", ln=True, align="L")

    pdf.ln(10)  # Add a new line

    # Add material type breakdown (Page 1)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Material Type Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Create a table for material type breakdown
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)  # Set background color for header row

    # Header row
    pdf.cell(70, 10, "Material Type", border=1, fill=True, align="C")
    pdf.cell(30, 10, "Count", border=1, fill=True, align="C")
    pdf.ln(10)  # Add a new line

    pdf.set_font("Arial", "", 10)
    material_types = []
    material_counts = []
    for material_type, group in grouped_data:
        # Calculate the required height for the cell based on the length of the material type name
        material_type_height = 10 + (pdf.get_string_width(material_type) // 55) * 10

        # Adjust the row height if the material type name exceeds the column width
        row_height = max(material_type_height, 10)  # Minimum row height is 10

        # Draw the material type cell with adjusted height
        pdf.cell(70, row_height, material_type, border=1, align="L")
        pdf.cell(30, row_height, str(len(group)), border=1, align="C")
        pdf.ln(row_height)  # Add a new line

        material_types.append(material_type)
        material_counts.append(len(group))

    pdf.add_page()  # Add a new page for Material Type Breakdown (Pie Chart) and Data Source and Filters

    # Add material type breakdown (Page 2)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Material Type Breakdown (Pie Chart):", ln=True, align="L")
    pdf.ln(5)  # Add space

    plt.figure(figsize=(13, 13))
    plt.pie(material_counts, labels=material_types, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the pie chart as an image in the Content folder
    chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/material_type_breakdown.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()

    # Add the image to the PDF report
    pdf.image(chart_path, x=50, y=pdf.get_y(), w=100)

    pdf.ln(100)  # Add space

    # Add a new pie chart for Total NET weight by material type
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Total NET Weight Breakdown by Material Type:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Calculate the total NET weight for each material type
    total_net_weights = [group['Total NET weight'].sum() for material_type, group in grouped_data]

    plt.figure(figsize=(17, 17))
    plt.pie(total_net_weights, labels=material_types, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the new pie chart as an image in the Content folder
    net_weight_chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/total_net_weight_breakdown.png"
    plt.savefig(net_weight_chart_path, bbox_inches='tight')
    plt.close()

    # Add the new image to the PDF report
    pdf.image(net_weight_chart_path, x=50, y=pdf.get_y(), w=100)

    # Add data source and filters information
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Source and Filters:", ln=True, align="L")
    pdf.ln(5)  # Add space

    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Data Source: {file_name}", ln=True, align="L")
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Filters Applied:", ln=True, align="L")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, "- SBM scope = True", ln=True, align="L")

    pdf.ln(10)  # Add space

    # Add data for each material type
    pdf.set_font("Arial", "B", 11)
    for material_type, group in grouped_data:
        pdf.set_font("Arial", "BI", 11)
        pdf.cell(0, 10, f"Material Type: {material_type}", ln=True, align="L")
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 10, f"Total NET weight (TON): {group['Total NET weight'].sum()}", ln=True, align="L")
        pdf.cell(0, 10, f"Average Unit Weight (Kg): {group['Unit Weight'].mean()}", ln=True, align="L")

        # Filter data by Quantity UOM METER and PCS
        filtered_data_meter = group[group['Quantity UOM'] == 'METER']
        filtered_data_pcs = group[group['Quantity UOM'] == 'PCE']

        # If data exists for METER and PCS, then print the corresponding total quantities
        if not filtered_data_meter.empty:
            pdf.cell(0, 10, f"Total QTY to commit (METER): {filtered_data_meter['Total QTY to commit'].sum()}", ln=True,
                     align="L")

        if not filtered_data_pcs.empty:
            pdf.cell(0, 10, f"Total QTY to commit (PCE): {filtered_data_pcs['Total QTY to commit'].sum()}", ln=True,
                     align="L")

        pdf.ln(5)  # Add a new line

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Piping"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"{project_number}_MTOPipingReport_SBMscope.pdf")
    pdf.output(pdf_path)


# YARD Scope
def generate_pdf_piping_yard_scope(data, project_number, file_name):
    # Convert Total NET weight to TON
    data['Total NET weight'] = data['Total NET weight'] / 1000

    # Group the data by material type
    grouped_data = data.groupby('Pipe Base Material')

    # Initialize PDF object
    pdf = FPDF()

    # Set up the PDF
    pdf.set_title("MTO Piping Analyze - YARD Scope")
    pdf.add_page()

    # Set font and font size for the title
    pdf.set_font("Arial", "B", 16)

    # Add title
    pdf.cell(0, 10, "Piping MTO Analyze - YARD Scope", ln=True, align="C")

    # Add date of analysis
    pdf.set_font("Arial", "I", 8)
    pdf.cell(0, 10, f"Date of Analysis: {date.today().strftime('%d-%b-%Y')}", ln=True, align="R")

    # Add summary statistics
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total Number of Records: {len(data)}", ln=True, align="L")
    pdf.cell(0, 10, f"Overall Total NET weight (TON): {data['Total NET weight'].sum()}", ln=True, align="L")
    pdf.cell(0, 10, f"Average Unit Weight (Kg): {data['Unit Weight'].mean()}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity Committed: {data['Total QTY to commit'].sum()}", ln=True, align="L")

    pdf.ln(10)  # Add a new line

    # Add material type breakdown (Page 1)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Material Type Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Create a table for material type breakdown
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)  # Set background color for header row

    # Header row
    pdf.cell(70, 10, "Material Type", border=1, fill=True, align="C")
    pdf.cell(30, 10, "Count", border=1, fill=True, align="C")
    pdf.ln(10)  # Add a new line

    pdf.set_font("Arial", "", 10)
    material_types = []
    material_counts = []
    for material_type, group in grouped_data:
        # Calculate the required height for the cell based on the length of the material type name
        material_type_height = 10 + (pdf.get_string_width(material_type) // 55) * 10

        # Adjust the row height if the material type name exceeds the column width
        row_height = max(material_type_height, 10)  # Minimum row height is 10

        # Draw the material type cell with adjusted height
        pdf.cell(70, row_height, material_type, border=1, align="L")
        pdf.cell(30, row_height, str(len(group)), border=1, align="C")
        pdf.ln(row_height)  # Add a new line

        material_types.append(material_type)
        material_counts.append(len(group))

    pdf.add_page()  # Add a new page for Material Type Breakdown (Pie Chart) and Data Source and Filters

    # Add material type breakdown (Page 2)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Material Type Breakdown (Pie Chart):", ln=True, align="L")
    pdf.ln(5)  # Add space

    plt.figure(figsize=(13, 13))
    plt.pie(material_counts, labels=material_types, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the pie chart as an image in the Content folder
    chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/material_type_breakdown_yard.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()

    # Add the image to the PDF report
    pdf.image(chart_path, x=50, y=pdf.get_y(), w=100)

    pdf.ln(100)  # Add space

    # Add a new pie chart for Total NET weight by material type
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Total NET Weight Breakdown by Material Type:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Calculate the total NET weight for each material type
    total_net_weights = [group['Total NET weight'].sum() for material_type, group in grouped_data]

    plt.figure(figsize=(17, 17))
    plt.pie(total_net_weights, labels=material_types, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the new pie chart as an image in the Content folder
    net_weight_chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/total_net_weight_breakdown_yard.png"
    plt.savefig(net_weight_chart_path, bbox_inches='tight')
    plt.close()

    # Add the new image to the PDF report
    pdf.image(net_weight_chart_path, x=50, y=pdf.get_y(), w=100)

    # Add data source and filters information
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Source and Filters:", ln=True, align="L")
    pdf.ln(5)  # Add space

    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Data Source: {file_name}", ln=True, align="L")
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Filters Applied:", ln=True, align="L")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, "- YARD scope = True", ln=True, align="L")

    pdf.ln(10)  # Add space
    total_qty_pce = 0
    total_qty_mtrs = 0

    # Add data for each material type
    pdf.set_font("Arial", "B", 11)
    for material_type, group in grouped_data:
        pdf.set_font("Arial", "BI", 11)
        pdf.cell(0, 10, f"Material Type: {material_type}", ln=True, align="L")
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 10, f"Total NET weight (TON): {group['Total NET weight'].sum()}", ln=True, align="L")
        pdf.cell(0, 10, f"Average Unit Weight (Kg): {group['Unit Weight'].mean()}", ln=True, align="L")

        # Filter data by Quantity UOM METER and PCS
        filtered_data_meter = group[group['Quantity UOM'] == 'METER']
        filtered_data_pcs = group[group['Quantity UOM'] == 'PCE']

        # If data exists for METER and PCS, then print the corresponding total quantities
        if not filtered_data_meter.empty:
            pdf.cell(0, 10, f"Total QTY to commit (METER): {filtered_data_meter['Total QTY to commit'].sum()}", ln=True,
                     align="L")
            total_qty_mtrs += filtered_data_meter['Total QTY to commit'].sum()

        if not filtered_data_pcs.empty:
            pdf.cell(0, 10, f"Total QTY to commit (PCE): {filtered_data_pcs['Total QTY to commit'].sum()}", ln=True,
                     align="L")
            total_qty_pce += filtered_data_pcs['Total QTY to commit'].sum()

        pdf.ln(5)  # Add a new line

    pdf.add_page()

    # Add insights
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 10, "Quantities Checks:", ln=True, align="L")
    pdf.ln(3)  # Add space
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10, f"Total Quantity in Meters: {total_qty_mtrs}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity in Pieces: {total_qty_pce}", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Piping"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"SO{project_number}_MTOPipingReport_YARDscope.pdf")
    pdf.output(pdf_path)


# -------------------------- Valve -----------------------------

# SBM Scope
def generate_pdf_valve_sbm_scope(data, project_number, file_name):
    # Convert Weight to TON
    data['Weight'] = data['Weight'] / 1000

    # Group the data by General Material Description
    grouped_data = data.groupby('General Material Description')

    # Initialize PDF object
    pdf = FPDF()

    # Set up the PDF
    pdf.set_title("Valve Equipment Analysis")
    pdf.add_page()

    # Set font and font size for the title
    pdf.set_font("Arial", "B", 16)

    # Add title
    pdf.cell(0, 10, "Valve MTO Analysis - SBM Scope", ln=True, align="C")

    # Add date of analysis
    pdf.set_font("Arial", "I", 12)
    pdf.cell(0, 10, f"Date of Analysis: {date.today().strftime('%d-%b-%Y')}", ln=True, align="R")

    # Add summary statistics
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total Number of Records: {len(data)}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity: {data['Quantity'].sum()}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Weight (TON): {data['Weight'].sum()}", ln=True, align="L")

    pdf.ln(10)  # Add a new line

    # Add material description breakdown
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Material Description Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Create a table for material description breakdown
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)  # Set background color for header row

    # Header row
    pdf.cell(80, 10, "Material Description", border=1, fill=True, align="C")
    pdf.cell(40, 10, "Total Quantity", border=1, fill=True, align="C")
    pdf.cell(40, 10, "Total Weight (TON)", border=1, fill=True, align="C")
    pdf.ln(10)  # Add a new line

    pdf.set_font("Arial", "", 10)
    material_descriptions = []
    total_quantities = []
    total_weights = []

    for material_desc, group in grouped_data:
        # Calculate the required height for the cell based on the length of the material description
        material_desc_height = 10 + (pdf.get_string_width(material_desc) // 55) * 10

        # Adjust the row height if the material description exceeds the column width
        row_height = max(material_desc_height, 10)  # Minimum row height is 10

        # Draw the material description cell with adjusted height
        pdf.cell(80, row_height, material_desc, border=1, align="L")
        pdf.cell(40, row_height, str(group['Quantity'].sum()), border=1, align="C")
        pdf.cell(40, row_height, str(group['Weight'].sum()), border=1, align="C")
        pdf.ln(row_height)  # Add a new line

        material_descriptions.append(material_desc)
        total_quantities.append(group['Quantity'].sum())
        total_weights.append(group['Weight'].sum())

    pdf.add_page()  # Add a new page for Material Description Breakdown (Pie Chart)

    # Add material description breakdown (Pie Chart)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Material Description Breakdown (Pie Chart):", ln=True, align="L")
    pdf.ln(5)  # Add space

    plt.figure(figsize=(13, 13))
    plt.pie(total_weights, labels=material_descriptions, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the pie chart as an image
    chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/material_description_breakdown_valve.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()

    # Add the image to the PDF report
    pdf.image(chart_path, x=50, y=pdf.get_y(), w=100)

    pdf.ln(100)  # Add space

    # Add data for each material description
    pdf.set_font("Arial", "B", 11)
    for material_desc, group in grouped_data:
        pdf.set_font("Arial", "BI", 11)
        pdf.cell(0, 10, f"Material Description: {material_desc}", ln=True, align="L")
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 10, f"Total Quantity: {group['Quantity'].sum()}", ln=True, align="L")
        pdf.cell(0, 10, f"Total Weight (TON): {group['Weight'].sum()}", ln=True, align="L")
        pdf.ln(5)  # Add a new line

    # Add data source and filters information
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Source and Filters:", ln=True, align="L")
    pdf.ln(5)  # Add space

    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Data Source: {file_name}", ln=True, align="L")
    pdf.cell(0, 10, "Filters Applied:", ln=True, align="L")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, "- SBM scope = True", ln=True, align="L")

    pdf.ln(10)  # Add space

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Valve"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"{project_number}_MTOValveReport_SBMscope.pdf")
    pdf.output(pdf_path)


# YARD Scope
def generate_pdf_valve_yard_scope(data, project_number, file_name):
    # Convert Weight to TON
    data['Weight'] = data['Weight'] / 1000

    # Group the data by General Material Description
    grouped_data = data.groupby('General Material Description')

    # Initialize PDF object
    pdf = FPDF()

    # Set up the PDF
    pdf.set_title("Valve MTO Analysis - YARD Scope")
    pdf.add_page()

    # Set font and font size for the title
    pdf.set_font("Arial", "B", 16)

    # Add title
    pdf.cell(0, 10, "Valve Equipment Analysis - YARD Scope", ln=True, align="C")

    # Add date of analysis
    pdf.set_font("Arial", "I", 12)
    pdf.cell(0, 10, f"Date of Analysis: {date.today().strftime('%d-%b-%Y')}", ln=True, align="R")

    # Add summary statistics
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total Number of Records: {len(data)}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity: {data['Quantity'].sum()}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Weight (TON): {data['Weight'].sum()}", ln=True, align="L")

    pdf.ln(10)  # Add a new line

    # Add material description breakdown
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Material Description Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Create a table for material description breakdown
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)  # Set background color for header row

    # Header row
    pdf.cell(80, 10, "Material Description", border=1, fill=True, align="C")
    pdf.cell(40, 10, "Total Quantity", border=1, fill=True, align="C")
    pdf.cell(40, 10, "Total Weight (TON)", border=1, fill=True, align="C")
    pdf.ln(10)  # Add a new line

    pdf.set_font("Arial", "", 10)
    material_descriptions = []
    total_quantities = []
    total_weights = []

    for material_desc, group in grouped_data:
        # Calculate the required height for the cell based on the length of the material description
        material_desc_height = 10 + (pdf.get_string_width(material_desc) // 55) * 10

        # Adjust the row height if the material description exceeds the column width
        row_height = max(material_desc_height, 10)  # Minimum row height is 10

        # Draw the material description cell with adjusted height
        pdf.cell(80, row_height, material_desc, border=1, align="L")
        pdf.cell(40, row_height, str(group['Quantity'].sum()), border=1, align="C")
        pdf.cell(40, row_height, str(group['Weight'].sum()), border=1, align="C")
        pdf.ln(row_height)  # Add a new line

        material_descriptions.append(material_desc)
        total_quantities.append(group['Quantity'].sum())
        total_weights.append(group['Weight'].sum())

    pdf.add_page()  # Add a new page for Material Description Breakdown (Pie Chart)

    # Add material description breakdown (Pie Chart)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Material Description Breakdown (Pie Chart):", ln=True, align="L")
    pdf.ln(5)  # Add space

    plt.figure(figsize=(13, 13))
    plt.pie(total_weights, labels=material_descriptions, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the pie chart as an image
    chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/material_description_breakdown_valve_yard.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()

    # Add the image to the PDF report
    pdf.image(chart_path, x=50, y=pdf.get_y(), w=100)

    pdf.ln(100)  # Add space

    # Add data for each material description
    pdf.set_font("Arial", "B", 11)
    for material_desc, group in grouped_data:
        pdf.set_font("Arial", "BI", 11)
        pdf.cell(0, 10, f"Material Description: {material_desc}", ln=True, align="L")
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 10, f"Total Quantity: {group['Quantity'].sum()}", ln=True, align="L")
        pdf.cell(0, 10, f"Total Weight (TON): {group['Weight'].sum()}", ln=True, align="L")
        pdf.ln(5)  # Add a new line

    # Add data source and filters information
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Source and Filters:", ln=True, align="L")
    pdf.ln(5)  # Add space

    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Data Source: {file_name}", ln=True, align="L")
    pdf.cell(0, 10, "Filters Applied:", ln=True, align="L")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, " - YARD scope = True", ln=True, align="L")

    pdf.ln(10)  # Add space

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Valve"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"{project_number}_MTOValveReport_YARDscope.pdf")
    pdf.output(pdf_path)


# --------------------------- Bolt --------------------------------

# SBM Scope
def generate_pdf_bolt_sbm_scope(data, project_number, file_name):
    # Group the data by Pipe Base Material
    grouped_data = data.groupby('Pipe Base Material')

    # Initialize PDF object
    pdf = FPDF()

    # Set up the PDF
    pdf.set_title("Bolt Equipment Analysis")
    pdf.add_page()

    # Set font and font size for the title
    pdf.set_font("Arial", "B", 16)

    # Add title
    pdf.cell(0, 10, "Bolt MTO Analysis - SBM Scope", ln=True, align="C")

    # Add date of analysis
    pdf.set_font("Arial", "I", 12)
    pdf.cell(0, 10, f"Date of Analysis: {date.today().strftime('%d-%b-%Y')}", ln=True, align="R")

    # Add summary statistics
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total Number of Records: {len(data)}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity Committed: {data['Total QTY to commit'].sum()}", ln=True, align="L")

    pdf.ln(10)  # Add a new line

    # Add base material breakdown (Page 1)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Base Material Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Create a table for base material breakdown
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)  # Set background color for header row

    # Header row
    pdf.cell(70, 10, "Base Material", border=1, fill=True, align="C")
    pdf.cell(30, 10, "Total Quantity", border=1, fill=True, align="C")
    pdf.ln(10)  # Add a new line

    pdf.set_font("Arial", "", 10)
    base_materials = []
    total_quantities = []

    for base_material, group in grouped_data:
        # Calculate the required height for the cell based on the length of the base material
        base_material_height = 10 + (pdf.get_string_width(base_material) // 55) * 10

        # Adjust the row height if the base material exceeds the column width
        row_height = max(base_material_height, 10)  # Minimum row height is 10

        # Draw the base material cell with adjusted height
        pdf.cell(70, row_height, base_material, border=1, align="L")
        pdf.cell(30, row_height, str(group['Total QTY to commit'].sum()), border=1, align="C")
        pdf.ln(row_height)  # Add a new line

        base_materials.append(base_material)
        total_quantities.append(group['Total QTY to commit'].sum())

    pdf.add_page()  # Add a new page for Base Material Breakdown (Pie Chart)

    # Add base material breakdown (Pie Chart)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Base Material Breakdown (Pie Chart):", ln=True, align="L")
    pdf.ln(5)  # Add space

    plt.figure(figsize=(13, 13))
    plt.pie(total_quantities, labels=base_materials, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the pie chart as an image
    chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/base_material_breakdown_bolt.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()

    # Add the image to the PDF report
    pdf.image(chart_path, x=50, y=pdf.get_y(), w=100)

    # Add data source and filters information
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Source and Filters:", ln=True, align="L")
    pdf.ln(5)  # Add space

    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Data Source: {file_name}", ln=True, align="L")
    pdf.cell(0, 10, "Filters Applied:", ln=True, align="L")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, "- SBM scope = True", ln=True, align="L")
    pdf.ln(10)  # Add space

    # Add data for each base material
    pdf.set_font("Arial", "B", 11)
    for base_material, group in grouped_data:
        pdf.set_font("Arial", "BI", 11)
        pdf.cell(0, 10, f"Base Material: {base_material}", ln=True, align="L")
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 10, f"Total Quantity Committed: {group['Total QTY to commit'].sum()}", ln=True, align="L")

        # Compare quantity confirmed and quantity committed
        qty_confirmed = group['Qty confirmed in design'].sum()
        qty_committed = group['Total QTY to commit'].sum()
        qty_difference = qty_confirmed - qty_committed

        pdf.cell(0, 10, f"Quantity Confirmed in Design: {qty_confirmed}", ln=True, align="L")
        pdf.cell(0, 10, f"Quantity Difference: {qty_difference}", ln=True, align="L")

        pdf.ln(5)  # Add a new line

    # Add Quantity UOM information
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 10, "Quantity UOM:", ln=True, align="L")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10, "All quantities are in PCE (Pieces)", ln=True, align="L")

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Bolt"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"{project_number}_MTOBoltReport_SBMscope.pdf")
    pdf.output(pdf_path)


# YARD Scope
def generate_pdf_bolt_yard_scope(data, project_number, file_name):
    # Group the data by Pipe Base Material
    grouped_data = data.groupby('Pipe Base Material')

    # Initialize PDF object
    pdf = FPDF()

    # Set up the PDF
    pdf.set_title("Bolt Equipment Analysis")
    pdf.add_page()

    # Set font and font size for the title
    pdf.set_font("Arial", "B", 16)

    # Add title
    pdf.cell(0, 10, "Bolt MTO Analysis - YARD Scope", ln=True, align="C")

    # Add date of analysis
    pdf.set_font("Arial", "I", 12)
    pdf.cell(0, 10, f"Date of Analysis: {date.today().strftime('%d-%b-%Y')}", ln=True, align="R")

    # Add summary statistics
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total Number of Records: {len(data)}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity Committed: {data['Total QTY to commit'].sum()}", ln=True, align="L")

    pdf.ln(10)  # Add a new line

    # Add base material breakdown (Page 1)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Base Material Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Create a table for base material breakdown
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)  # Set background color for header row

    # Header row
    pdf.cell(70, 10, "Base Material", border=1, fill=True, align="C")
    pdf.cell(30, 10, "Total Quantity", border=1, fill=True, align="C")
    pdf.ln(10)  # Add a new line

    pdf.set_font("Arial", "", 10)
    base_materials = []
    total_quantities = []

    for base_material, group in grouped_data:
        # Calculate the required height for the cell based on the length of the base material
        base_material_height = 10 + (pdf.get_string_width(base_material) // 55) * 10

        # Adjust the row height if the base material exceeds the column width
        row_height = max(base_material_height, 10)  # Minimum row height is 10

        # Draw the base material cell with adjusted height
        pdf.cell(70, row_height, base_material, border=1, align="L")
        pdf.cell(30, row_height, str(group['Total QTY to commit'].sum()), border=1, align="C")
        pdf.ln(row_height)  # Add a new line

        base_materials.append(base_material)
        total_quantities.append(group['Total QTY to commit'].sum())

    pdf.add_page()  # Add a new page for Base Material Breakdown (Pie Chart)

    # Add base material breakdown (Pie Chart)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Base Material Breakdown (Pie Chart):", ln=True, align="L")
    pdf.ln(5)  # Add space

    plt.figure(figsize=(13, 13))
    plt.pie(total_quantities, labels=base_materials, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the pie chart as an image
    chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/base_material_breakdown_bolt_yard.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()

    # Add the image to the PDF report
    pdf.image(chart_path, x=50, y=pdf.get_y(), w=100)

    # Add data source and filters information
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Source and Filters:", ln=True, align="L")
    pdf.ln(5)  # Add space

    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Data Source: {file_name}", ln=True, align="L")
    pdf.cell(0, 10, "Filters Applied:", ln=True, align="L")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, "- YARD scope = True", ln=True, align="L")
    pdf.ln(10)  # Add space

    # Add data for each base material
    pdf.set_font("Arial", "B", 11)
    for base_material, group in grouped_data:
        pdf.set_font("Arial", "BI", 11)
        pdf.cell(0, 10, f"Base Material: {base_material}", ln=True, align="L")
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 10, f"Total Quantity Committed: {group['Total QTY to commit'].sum()}", ln=True, align="L")

        # Compare quantity confirmed and quantity committed
        qty_confirmed = group['Qty confirmed in design'].sum()
        qty_committed = group['Total QTY to commit'].sum()
        qty_difference = qty_confirmed - qty_committed

        pdf.cell(0, 10, f"Quantity Confirmed in Design: {qty_confirmed}", ln=True, align="L")
        pdf.cell(0, 10, f"Quantity Difference: {qty_difference}", ln=True, align="L")

        pdf.ln(5)  # Add a new line

    # Add Quantity UOM information
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 10, "Quantity UOM:", ln=True, align="L")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10, "All quantities are in PCE (Pieces)", ln=True, align="L")

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Bolt"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"{project_number}_MTOBoltReport_YARDscope.pdf")
    pdf.output(pdf_path)


# ----------------------- STRUCTURE ------------------------------


def generate_pdf_structure_sbm_scope1(data, project_number, file_name):
    # Create a dictionary mapping product codes to short names
    code_to_short = {
        'BLK.STR.PRI.PLT': 'PRI.PLT',
        'BLK.STR.PRI.PRO': 'PRI.PRO',
        'BLK.STR.PRI.TBR': 'PRI.TBR',
        'BLK.STR.SEC.PLT': 'SEC.PLT',
        'BLK.STR.SEC.PRO': 'SEC.PRO',
        'BLK.STR.SEC.TBR': 'SEC.TBR',
        'BLK.STR.SPS.PLT': 'SPS.PLT',
        'BLK.STR.SPS.TBR': 'SPS.TBR',
        'BLK.STR.TER.PLT': 'TER.PLT',
        'BLK.STR.TER.PRO': 'TER.PRO',
        'BLK.STR.TER.SNT': 'TER.SNT',
    }

    # Create a dictionary mapping short names to full names
    short_to_full = {
        'PRI.PLT': 'Primary Steel, Plate',
        'PRI.PRO': 'Primary Steel, Profiles',
        'PRI.TBR': 'Primary Steel, Tubulars',
        'SEC.PLT': 'Secondary Steel, Plate',
        'SEC.PRO': 'Secondary Steel, Profiles',
        'SEC.TBR': 'Secondary Steel, Tubulars',
        'SPS.PLT': 'Special Steel, Plate',
        'SPS.TBR': 'Special Steel, Tubulars',
        'TER.PLT': 'Tertiary Steel, Plate',
        'TER.PRO': 'Tertiary Steel, Profiles',
        'TER.SNT': 'Tertiary Steel, Subassembly & The Like',
    }

    # Convert Total NET weight to TON
    data['Total NET weight'] = data['Total NET weight'] / 1000
    df = data.copy()

    # Group the data by product code and quantity UOM
    grouped_data = data.groupby(['Product Code', 'Quantity UOM'])

    # Initialize PDF object
    pdf = FPDF()

    # Set up the PDF
    pdf.set_title("Structure Equipment Data Analysis")
    pdf.add_page()

    # Set font and font size for the title
    pdf.set_font("Arial", "B", 16)

    # Add title
    pdf.cell(0, 10, "Structure Equipment Data Analysis", ln=True, align="C")

    # Add date of analysis
    pdf.set_font("Arial", "I", 12)
    pdf.cell(0, 10, f"Date of Analysis: {date.today().strftime('%d-%b-%Y')}", ln=True, align="R")

    # Add summary statistics
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total Number of Records: {len(data)}", ln=True, align="L")
    pdf.cell(0, 10, f"Overall Total NET weight (TON): {data['Total NET weight'].sum()}", ln=True, align="L")
    pdf.cell(0, 10, f"Average Unit Weight (Kg): {data['Unit Weight'].mean()}", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity Committed: {data['Total QTY to commit'].sum()}", ln=True, align="L")

    pdf.ln(10)  # Add a new line

    # Add data breakdown
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Create a table for data breakdown
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)  # Set background color for header row

    # Header row
    pdf.cell(70, 10, "Product Code", border=1, fill=True, align="C")
    pdf.cell(30, 10, "Quantity UOM", border=1, fill=True, align="C")
    pdf.cell(40, 10, "Total QTY to commit", border=1, fill=True, align="C")
    pdf.cell(40, 10, "Total NET weight (TON)", border=1, fill=True, align="C")
    pdf.ln(10)  # Add a new line

    pdf.set_font("Arial", "", 10)
    for (product_code, quantity_uom), group in grouped_data:
        # Replace the product code with its short name
        product_code_short = code_to_short[product_code]

        # Draw the cells for data breakdown
        pdf.cell(70, 10, product_code_short, border=1, align="L")
        pdf.cell(30, 10, quantity_uom, border=1, align="C")
        pdf.cell(40, 10, str(group['Total QTY to commit'].sum()), border=1, align="R")
        pdf.cell(40, 10, f"{group['Total NET weight'].sum():.2f}", border=1, align="R")
        pdf.ln(10)  # Add a new line

    pdf.add_page()  # Add a new page for Graphics and Images

    # Add graphics and images
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Graphics and Images:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Generate figures and add them to the PDF report
    for column in ['Total QTY to commit', 'Total NET weight', 'Unit Weight', 'Thickness', 'Wastage Quantity',
                   'Quantity Including Wastage', 'Total Gross Weight']:
        # Create figure
        fig, ax = plt.subplots(figsize=(12, 5))

        # Plot the selected data
        barplot = sns.barplot(x='Product Code', y=column, hue='Quantity UOM', data=df, ax=ax)

        ax.set_title(f'{column} per Material Type and Quantity UOM')
        ax.set_xlabel('Material Type')
        ax.set_ylabel(column)

        graphics_dir = "../Data Pool/DCT Process Results/PDF Reports/Content/"

        # Add value labels on the bars
        for i, bar in enumerate(barplot.patches):
            if np.isfinite(bar.get_height()):
                barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                             f'{bar.get_height():.2f}',
                             ha='center', va='bottom',
                             fontsize=8, color='black')

        # Save the figure
        fig_path = os.path.join(graphics_dir, f"MP{project_number}_Structure_{column}_Graphic.png")
        plt.savefig(fig_path, bbox_inches='tight')
        plt.close(fig)

        # Add the figure image to the PDF report
        pdf.image(fig_path, x=10, y=pdf.get_y(), w=190)
        pdf.ln(120)  # Add space

    # Add data source and filters information
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Source and Filters:", ln=True, align="L")
    pdf.ln(5)  # Add space

    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Data Source: {file_name}", ln=True, align="L")
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Filters Applied:", ln=True, align="L")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, "- N/A", ln=True, align="L")

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Structure"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"{project_number}_StructureDataReport.pdf")
    pdf.output(pdf_path)


code_to_short = {
    'BLK.STR.PRI.PLT': 'Primary Steel, Plate',
    'BLK.STR.PRI.PRO': 'Primary Steel, Profiles',
    'BLK.STR.PRI.TBR': 'Primary Steel, Tubulars',
    'BLK.STR.SEC.PLT': 'Secondary Steel, Plate',
    'BLK.STR.SEC.PRO': 'Secondary Steel, Profiles',
    'BLK.STR.SEC.TBR': 'Secondary Steel, Tubulars',
    'BLK.STR.SPS.PLT': 'Special Steel, Plate',
    'BLK.STR.SPS.TBR': 'Special Steel, Tubulars',
    'BLK.STR.TER.PLT': 'Tertiary Steel, Plate',
    'BLK.STR.TER.PRO': 'Tertiary Steel, Profiles',
    'BLK.STR.TER.SNT': 'Tertiary Steel, Subassembly & The Like',
}

short_to_full = {
    'PRI.PLT': 'Primary Steel, Plate',
    'PRI.PRO': 'Primary Steel, Profiles',
    'PRI.TBR': 'Primary Steel, Tubulars',
    'SEC.PLT': 'Secondary Steel, Plate',
    'SEC.PRO': 'Secondary Steel, Profiles',
    'SEC.TBR': 'Secondary Steel, Tubulars',
    'SPS.PLT': 'Special Steel, Plate',
    'SPS.TBR': 'Special Steel, Tubulars',
    'TER.PLT': 'Tertiary Steel, Plate',
    'TER.PRO': 'Tertiary Steel, Profiles',
    'TER.SNT': 'Tertiary Steel, Subassembly & The Like',
}


def generate_pdf_structure_sbm_scope(dataframe, project_number):
    # Convert Total NET weight to TON
    dataframe['Total NET weight'] = dataframe['Total NET weight'] / 1000

    # Filter the dataframe based on the project number
    filtered_df = dataframe.copy()

    # Apply code_to_short mapping to Product Code column
    filtered_df['Product Code'] = filtered_df['Product Code'].map(code_to_short)

    # Group the filtered data by material type
    grouped_data = filtered_df.groupby('Product Code')

    # Initialize PDF object
    pdf = FPDF()

    # Set up the PDF
    pdf.set_title("Equipment Analysis Report")
    pdf.add_page()

    # Set font and font size for the title
    pdf.set_font("Arial", "B", 16)

    # Add title
    pdf.ln(20)
    pdf.cell(0, 10, f"MP17033 - Unity Structure MTO Analyze", ln=True, align="C")

    # Add date of analysis
    pdf.ln(3)
    pdf.set_font("Arial", "I", 8)
    pdf.cell(0, 10, f"Date of Analysis: {date.today().strftime('%d-%b-%Y')}", ln=True, align="R")

    pdf.ln(5)

    # Add logo
    logo_path = "../Data Pool/DCT Process Results/PDF Reports/Content/SBM.png"
    pdf.image(logo_path, x=10, y=10, w=30)

    # Add summary statistics
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total Number of Records: {len(filtered_df)}", ln=True, align="L")
    pdf.cell(0, 10, f"Overall Total Gross Weight: {filtered_df['Total Gross Weight'].sum():.4f} (TONs)", ln=True,
             align="L")
    pdf.cell(0, 10, f"Average Unit Weight: {filtered_df['Unit Weight'].mean():.4f} (Kg)", ln=True, align="L")
    pdf.cell(0, 10, f"Total Quantity Committed: {filtered_df['Total QTY to commit'].sum():.4f}", ln=True, align="L")

    pdf.ln(10)  # Add a new line

    # Add Material Type Breakdown (Page 1)
    pdf.set_font("Arial", "B", 13)
    pdf.cell(0, 10, "Material Type Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Create a table for Material Type Breakdown
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)  # Set background color for header row

    # Header row
    pdf.cell(70, 10, "Material Type", border=1, fill=True, align="C")
    pdf.cell(30, 10, "Count", border=1, fill=True, align="C")
    pdf.ln(10)  # Add a new line

    pdf.set_font("Arial", "", 9)
    material_types = []
    material_counts = []
    for material_type, group in grouped_data:
        # Calculate the required height for the cell based on the length of the material type name
        material_type_height = 10 + (pdf.get_string_width(material_type) // 55) * 10

        # Adjust the row height if the material type name exceeds the column width
        row_height = max(material_type_height, 10)  # Minimum row height is 10

        # Draw the material type cell with adjusted height
        pdf.cell(70, row_height, material_type, border=1, align="L")
        pdf.cell(30, row_height, str(len(group)), border=1, align="C")
        pdf.ln(row_height)  # Add a new line

        material_types.append(material_type)
        material_counts.append(len(group))

    pdf.add_page()  # Add a new page for Material Type Breakdown (Pie Chart) and Data Source and Filters

    # Add Material Type Breakdown (Page 2)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Material Type Breakdown (Pie Chart):", ln=True, align="L")
    pdf.ln(5)  # Add space

    plt.figure(figsize=(13, 13))
    plt.pie(material_counts, labels=material_types, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the pie chart as an image in the Content folder
    chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/material_type_breakdown.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()

    # Add the image to the PDF report
    pdf.image(chart_path, x=50, y=pdf.get_y(), w=100)

    pdf.ln(100)  # Add space

    # Add a new pie chart for Total Gross Weight by material type
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Total Gross Weight Breakdown by Material Type:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Calculate the total Gross weight for each material type
    total_gross_weights = [group['Total Gross Weight'].sum() for material_type, group in grouped_data]

    plt.figure(figsize=(17, 17))
    plt.pie(total_gross_weights, labels=material_types, autopct='%1.1f%%')
    plt.axis('equal')

    # Save the new pie chart as an image in the Content folder
    gross_weight_chart_path = "../Data Pool/DCT Process Results/PDF Reports/Content/total_gross_weight_breakdown.png"
    plt.savefig(gross_weight_chart_path, bbox_inches='tight')
    plt.close()

    # Add the new image to the PDF report
    pdf.image(gross_weight_chart_path, x=50, y=pdf.get_y(), w=100)

    # Add Data Source and Filters information
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Source and Filters:", ln=True, align="L")
    pdf.ln(5)  # Add space

    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, "Data Source: [Specify data source]", ln=True, align="L")
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Filters Applied:", ln=True, align="L")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"- Project: {project_number}", ln=True, align="L")

    pdf.ln(10)  # Add space

    # Add Data Breakdown and Insights
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Data Breakdown:", ln=True, align="L")
    pdf.ln(5)  # Add space

    # Add breakdown for each material type
    pdf.set_font("Arial", "B", 11)
    for material_type, group in grouped_data:
        full_name = short_to_full.get(material_type, material_type)
        pdf.set_font("Arial", "BI", 11)
        pdf.cell(0, 10, f"Material Type: {full_name}", ln=True, align="L")
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 10, f"Total Gross Weight: {group['Total Gross Weight'].sum():.4f}", ln=True, align="L")
        pdf.cell(0, 10, f"Total NET Weight: {group['Total NET weight'].sum():.4f} (TONS)", ln=True, align="L")
        pdf.cell(0, 10, f"Total Quantity including Wastage: {group['Quantity Including Wastage'].sum():.4f}", ln=True,
                 align="L")
        pdf.cell(0, 10, f"Total Thickness: {group['Thickness'].sum():.4f}", ln=True, align="L")
        pdf.cell(0, 10, f"Average Thickness: {group['Thickness'].mean():.4f}", ln=True, align="L")
        pdf.cell(0, 10, f"Total QTY to commit: {group['Total QTY to commit'].sum():.4f}", ln=True, align="L")
        pdf.cell(0, 10, f"Wastage Quantity: {group['Wastage Quantity'].sum():.4f}", ln=True, align="L")
        pdf.ln(5)  # Add a new line

    # Calculate the percentage of Primary Steel, Plate
    primary_plate_count = len(grouped_data.get_group('Primary Steel, Plate'))
    total_records = len(filtered_df)
    primary_plate_percentage = (primary_plate_count / total_records) * 100

    # Calculate the percentage of Secondary Steel, Profiles
    sec_plate_count = len(grouped_data.get_group('Secondary Steel, Profiles'))
    total_records_sec = len(filtered_df)
    sec_plate_percentage = (sec_plate_count / total_records_sec) * 100

    # Calculate the percentage of Special Steel, Tubulars
    spc_plate_count = len(grouped_data.get_group('Special Steel, Tubulars'))
    total_records_spc = len(filtered_df)
    spc_plate_percentage = (spc_plate_count / total_records_spc) * 100

    # Add Insights
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 10, "Insights:", ln=True, align="L")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10, "The analysis of equipment data for Unity Project provides the following insights:", ln=True,
             align="L")
    pdf.cell(0, 10,
             f"- Primary Steel, Plate is the most common material type, accounting for {primary_plate_percentage:.4f} % of the records.",
             ln=True, align="L")
    pdf.cell(0, 10,
             f"- Secondary Steel, Profiles have a relatively high representation, accounting for {sec_plate_percentage:.4f} % of the records.",
             ln=True, align="L")
    pdf.cell(0, 10,
             f"- Special Steel, Tubular have the lowest representation, accounting for {spc_plate_percentage:.4f} % of the records.",
             ln=True, align="L")
    pdf.cell(0, 10, f"- The average unit weight across all materials is {filtered_df['Unit Weight'].mean():.4f} kg.",
             ln=True, align="L")
    pdf.cell(0, 10, f"- The overall total gross weight is {filtered_df['Total Gross Weight'].sum():.4f} tons.", ln=True,
             align="L")
    pdf.cell(0, 10, f"- The total quantity committed is {filtered_df['Total QTY to commit'].sum():.4f}.", ln=True,
             align="L")
    pdf.cell(0, 10,
             "- The data breakdown by material type and the corresponding quantities provide a detailed overview of the equipment distribution.",
             ln=True, align="L")

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Structure"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"{project_number}_MTOStructureReport_SBMscope.pdf")
    pdf.output(pdf_path)


# ----------------------- MTO Analyze ------------------------------


# Complete MTO PDF file analyze
def generate_unity_complete_analyze_process_pdf(piping_data, piping_sbm_data, piping_data_yard, total_qty_commit_pieces,
                                          total_qty_commit_m, total_piping_net_weight, total_piping_sbm_net_weight,
                                          total_piping_yard_net_weight, total_qty_commit_pieces_sbm,
                                          total_qty_commit_pieces_yard, total_qty_commit_m_sbm, total_qty_commit_m_yard,
                                          valve_data_sbm, valve_data_yard, total_valve_weight, total_sbm_valve_weight,
                                          total_yard_valve_weight,
                                          bolt_data_total_qty_commit, bolt_sbm_data_total_qty_commit,
                                          bolt_yard_data_total_qty_commit, structure_totals_m2, structure_totals_m,
                                          structure_totals_pcs,
                                          total_matched_tags_pip, total_unmatched_tags_pip, total_surplus_tags_pip,
                                          total_weight_pip, total_quantity_by_uom_pip, overall_cost_pip,
                                          total_cost_by_material_pip, unique_cost_object_ids_pip,
                                          total_surplus_cost_pip, unique_surplus_cost_object_ids_pip,
                                          total_spcpip_data_weight, total_spcpip_data_qty, total_spcpip_sbm_data_weight,
                                          total_spcpip_sbm_data_qty, total_spcpip_yard_data_qty,
                                          total_spcpip_yard_data_weight,
                                          total_quantity_vlv, overall_cost_vlv, cost_by_general_description_vlv,
                                          total_po_quantity_blt, overall_cost_blt, cost_by_pipe_base_material_blt,
                                          missing_product_codes_blt, structure_total_gross_weight,
                                          structure_total_wastage, structure_total_qty_pcs, structure_total_qty_m2,
                                          structure_total_qty_m, total_matched_tags_spc, total_unmatched_tags_spc,
                                          total_quantity_by_uom_spc, total_cost_spc, po_list_spc,
                                          project_total_cost_and_hours, total_surplus_plus_tags,
                                          total_po_quantity_piece, total_po_quantity_meter, unmatch_pip_total_weight, unmatch_pip_total_qty, unmatch_pip_avg_unit_weight):

    # Title and sub title functions
    def add_section_title(title):
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, title, ln=True)
        pdf.set_font("Arial", "", 10)

    def add_section_sub_title(subtitle):
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, subtitle, ln=True)
        pdf.set_font("Arial", "", 10)

    def add_table(header, data):
        pdf.set_font("Arial", "B", 10)
        for item in header:
            pdf.cell(40, 10, item, border=1, align="C")
        pdf.ln()
        pdf.set_font("Arial", "", 10)
        for row in data:
            for item in row:
                pdf.cell(40, 10, str(item), border=1, align="C")
            pdf.ln()

    # Function to add an image with a title below it
    def add_image_with_legend(pdf, folder_path, image_name_start, legend, x=20, y=None, w=0, h=0):
        # Searching for image in the folder
        matching_images = [img for img in os.listdir(folder_path) if img.startswith(image_name_start)]
        if not matching_images:
            print(f"No images found that start with '{image_name_start}' in '{folder_path}'")
            return
        image_path = os.path.join(folder_path, matching_images[0])

        # If y is not provided, set it to the current Y position of the pdf
        if y is None:
            y = pdf.get_y()

        # Add the image first
        pdf.image(image_path, x=x, y=y, w=w, h=h)

        # If h is not provided, calculate the height of the image based on its aspect ratio
        if not h:
            # Get image dimensions (in pixels)
            from PIL import Image
            with Image.open(image_path) as img:
                width_px, height_px = img.size

            # Calculate the aspect ratio of the image
            aspect_ratio = height_px / width_px

            # Calculate the height in the PDF based on the provided width w and the image's aspect ratio
            h = w * aspect_ratio
            gap = h
        else:
            gap = h

        # Set position for the legend, centered below the image
        # Adjust the position for the legend, slightly to the left of center
        offset = w * 0.2  # 20% offset to the left from the center
        pdf.set_xy(x + w / 2 - pdf.get_string_width(legend) / 2 - offset, y + gap + 1)

        # Set font for the legend and print it
        pdf.set_font('Arial', '', 8)
        pdf.cell(0, 10, legend, ln=True, align="C")

    # Function to convert weight to tons
    def convert_to_tons(weight_kg):
        return weight_kg / 1000

    # Function to convert cost to thousands
    def convert_to_thousands(cost):
        return cost / 1000

    # Initialize PDF object
    pdf = FPDF()

    # Set up the PDF
    pdf.set_title("MP17033 - UNITY MTO Material and Cost Analyze")
    pdf.add_page()

    # Set font and font size for the title
    pdf.set_font("Arial", "B", 17)

    # Add logo
    logo_path = "../Data Pool/DCT Process Results/PDF Reports/Content/SBM.png"
    pdf.image(logo_path, x=10, y=10, w=30)

    # Add title
    pdf.ln(19)
    pdf.cell(0, 10, "MP17033 Unity - Data Collection and Transformation Report", ln=True, align="C")

    # Add date of analysis
    pdf.ln(1)
    pdf.set_font("Arial", "I", 6)
    pdf.cell(0, 10, f"Date of Analysis: {date.today().strftime('%d-%b-%Y')}", ln=True, align="R")

    pdf.ln(1)

    image_path = "../Data Pool/DCT Process Results/PDF Reports/Content/LizaUnity1.jpg"
    # Get the original image dimensions
    img_w = 1178
    img_h = 400

    # Determine the desired width for the image in the PDF
    desired_width = 175

    # Calculate the proportional height based on the desired width
    proportional_height = img_h * desired_width / img_w

    # Add the image with the desired dimensions
    pdf.image(image_path, x=18, y=55, w=desired_width, h=proportional_height)

    pdf.ln(66)

    # Introduction
    add_section_title("Introduction")
    pdf.multi_cell(0, 8,
                   "This report provides a comprehensive analysis of the cost associated with different Equipment in the Unity FPSO Project. The analysis focuses on Piping, Special Piping, Structure, Valves and Bolts within the realms of SBM Scope and YARD Scope. "
                   "The report includes an overview of all scopes and equipment involved, followed by detailed insights for each equipment. "
                   "The report elucidates the total expenditure, alongside detailed cost breakdowns for each equipment category, metrics on cost per kilogram, and cost per supplier.",
                   align="J")

    pdf.ln(3)

    # Overview of Scopes and Materials
    add_section_title("Overview of Scopes and Materials")
    pdf.cell(0, 9, "This analysis of the Unity Project consists of two different scopes, including:", ln=True)
    # SBM Scope
    pdf.cell(8)
    pdf.cell(0, 10, "a. SBM Scope", ln=True)
    pdf.cell(18)
    pdf.cell(0, 7, "- Piping", ln=True)
    pdf.cell(18)
    pdf.cell(0, 7, "- Special Piping", ln=True)
    pdf.cell(18)
    pdf.cell(0, 7, "- Valve", ln=True)
    pdf.cell(18)
    pdf.cell(0, 7, "- Bolt", ln=True)

    # YARD Scope
    pdf.cell(8)
    pdf.cell(0, 10, "b. YARD Scope", ln=True)
    pdf.cell(18)
    pdf.cell(0, 7, "- Piping", ln=True)
    pdf.cell(18)
    pdf.cell(0, 7, "- Special Piping", ln=True)
    pdf.cell(18)
    pdf.cell(0, 7, "- Structure", ln=True)
    pdf.cell(18)
    pdf.cell(0, 7, "- Valve", ln=True)
    pdf.cell(18)
    pdf.cell(0, 7, "- Bolt", ln=True)

    pdf.add_page()

    # Calculations:

    total_dct_weight = convert_to_tons((total_weight_pip + total_valve_weight + structure_total_gross_weight + total_spcpip_data_weight))
    total_dct_mto_cost = convert_to_thousands(total_cost_spc + overall_cost_pip + overall_cost_vlv + overall_cost_blt)
    total_project_cost = convert_to_thousands(project_total_cost_and_hours[0])
    total_project_hours = project_total_cost_and_hours[1]
    total_dct_pieces = bolt_data_total_qty_commit + total_qty_commit_pieces + total_spcpip_data_qty + total_quantity_vlv + structure_total_qty_pcs
    total_dct_meters_yard = structure_total_qty_m2 + structure_total_qty_m
    total_dct_sbm_pieces = total_spcpip_sbm_data_qty + bolt_sbm_data_total_qty_commit + total_qty_commit_pieces_sbm + total_quantity_vlv
    total_dct_yard_pieces = total_qty_commit_pieces_yard + total_spcpip_yard_data_qty + bolt_yard_data_total_qty_commit + structure_total_qty_pcs
    total_dct_sbm_weight = convert_to_tons(total_piping_sbm_net_weight + total_sbm_valve_weight + total_spcpip_sbm_data_weight)
    total_dct_yard_weight = convert_to_tons(total_yard_valve_weight + total_piping_yard_net_weight + structure_total_gross_weight + total_spcpip_yard_data_weight)
    total_dct_meters = structure_total_qty_m2 + structure_total_qty_m + total_qty_commit_m
    total_dct_piping_surplus = total_surplus_tags_pip
    total_dct_piping_surplus_plus = total_surplus_plus_tags
    unmatch_pip_perc = (total_unmatched_tags_pip / (total_unmatched_tags_pip + total_matched_tags_pip)) * 100

    total_dct_piping_surplus_cost = convert_to_thousands(abs(total_surplus_cost_pip))
    total_structure_wastage = structure_total_wastage
    total_dct_piping_weight = convert_to_tons(total_piping_net_weight)
    total_dct_valve_weight = convert_to_tons(total_valve_weight)
    total_dct_structure_weight = structure_total_gross_weight
    total_dct_scp_piping_weight = convert_to_tons(total_spcpip_data_weight)
    total_dct_po_quantity_piece = total_po_quantity_piece + total_quantity_by_uom_spc + total_po_quantity_blt
    total_dct_po_quantity_meter = total_po_quantity_meter

    pdf.ln(5)
    add_section_title("Data Collection and Transformation: Enhanced Project Analysis Report")
    pdf.ln(5)
    add_section_sub_title("Overall Summary")
    pdf.ln(3)
    pdf.multi_cell(0, 8,
                   "The Unity Project FPSO stands as a key project for our company. It signifies deep operational and financial implications."
                   f" Our in-depth analysis provides insights into the project's scope and complexities. Upon examining, the total material weight for the project is approximately {total_dct_weight:.3f} metric tons."
                   f" This includes various components from small fittings to major structural parts. The Material Take Off (MTO) shows an equipment cost of ${total_dct_mto_cost:.3f} thousands USD."
                   f" Including all aspects of the Unity Project FPSO, the total financial implication is about ${total_project_cost:.2f} thousands UDS (a sum of all purchased order)."
                   f" Manpower and time are crucial. We obtained approximately a total of {total_project_hours:.2f} hours for this project alone."
                   f" As Unity project is at its end we had a good amount of data to work with. The project demands {total_dct_pieces:.0f} individual material pieces, highlighting its intricacy."
                   f" About {total_dct_meters:.3f} meters of material is needed, with {total_dct_meters_yard:.1f} meters sourced from the YARD.", align="J")

    pdf.ln(7)

    # Scopes Breakdown
    add_section_sub_title("Detailed Scope Analysis")
    pdf.ln(3)
    pdf.multi_cell(0, 10, f"As mentioned above, in our assessment, we've distinguished between two primary scopes: SBM and YARD Scope."
                          f" Under the SBM scope, there's a total of {total_dct_sbm_pieces:.0f} pieces, markedly higher when compared to the YARD scope, which totals {total_dct_yard_pieces:.0f} pieces."
                          f" When considering weight, materials in the SBM scope amass to {total_dct_sbm_weight:.3f} tons. On the other hand, YARD scope materials have a cumulative weight of {total_dct_yard_weight:.3f} tons."
                          f" These figures underline the material distribution and weight disparities between the two pivotal project sectors.", align="J")

    pdf.ln(7)

    # Material Types Breakdown
    add_section_sub_title("Material Types Breakdown")
    pdf.ln(3)
    pdf.multi_cell(0, 10, "The weight distribution across various material types reveals key insights into the project's construction."
                          f" Piping emerges as the dominant equipment, accounting for a notable {total_dct_piping_weight:.3f} tons. Following closely are the valves, which contribute {total_dct_valve_weight:.3f} tons, making them the second most weighty category in the project."
                          f" Structural elements, pivotal to the project, register a total weight of {total_dct_structure_weight:.1f} tons. Not to be overlooked, special piping materials, a distinct category in our inventory, weigh {total_dct_scp_piping_weight:.3f} tons."
                          f" Such a breakdown accentuates the emphasis of particular material types, echoing the project's functional requirements and design choices.", align="J")

    pdf.add_page()

    # Surplus and Wastage
    pdf.ln(5)
    add_section_sub_title("Understanding Surplus and Wastage")
    pdf.ln(3)
    pdf.multi_cell(0, 10, f"In the evaluation of the project's material management, we identify two essential components: surplus and wastage. The surplus primarily manifests in piping, where an additional {total_dct_piping_surplus} pieces are noted. These surplus materials represent an associated cost of ${total_dct_piping_surplus_cost} thousands of dollars, highlighting the need for efficient utilization."
                          f"Concurrently, we have identified wastage within structural equipments. Approximately {total_structure_wastage:.1f} structural equipment wastage have been observed. This insight emphasizes the significance of meticulous material planning and allocation to minimize waste and align with budgetary constraints.", align="J")

    pdf.ln(7)

    # Purchase Order (PO) Metrics
    add_section_sub_title("Purchase Order Metrics")
    pdf.ln(3)
    pdf.multi_cell(0, 10, f"The overall total pieces of piping, special piping and bolt equipments analyzed from the PO's placed in NADIA are {total_dct_po_quantity_piece:.0f}, and the total meters mainly from piping are {total_dct_po_quantity_meter:.0f}.", align="")

    pdf.add_page()
    pdf.ln(1)
    #Equipment Details
    add_section_title("SBM SCOPE - Breakdown")
    pdf.ln(1)
    add_section_sub_title("Piping Specification - SBM Scope")
    pdf.cell(0, 5, f"The cumulative weight of the piping stands at {convert_to_tons(total_piping_sbm_net_weight):.2f} Tons, associated with a financial outlay of {convert_to_thousands(overall_cost_pip):.2f} thousand USD.", ln=True)
    pdf.cell(0, 5, f"The total number of piping pieces amounts to {total_qty_commit_pieces_sbm:.0f}, covering a material length of {total_qty_commit_m_sbm:.2f} meters.", ln=True)
    pdf.multi_cell(0, 5, f"Comparing the MTO file with Ecosys Data we have about {unmatch_pip_perc: .0f}% of Piping without a cost associated, weighting {convert_to_tons(unmatch_pip_total_weight): .0f} tons from a total of {unmatch_pip_total_qty} pieces.", align="J")

    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Piping", "MP17033_PipingTotal_Net_Weight_Per_Material_", "1 - Total Piping Weight per different Materials.", x=30, w=150)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Piping", "MP17033_PipingCostKG_Material_", "2 - Piping Cost per Kilo of each Material Type.", x=30, w=150)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Piping", "MP17033_PipingTotal_Cost_Material_", "3 - Piping Total Cost representation per each material type.", x=30, w=150)

    pdf.add_page()
    pdf.ln(5)

    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Piping", "MP17033_Piping_CostPerSupplier_Graphic_", "4 - Graphic representation of the Piping cost by different suppliers.", x=30, w=150)

    pdf.ln(5)

    add_section_sub_title("Bolt Specification - SBM Scope")
    pdf.cell(0, 7, f"The total count of bolts is approximately {bolt_sbm_data_total_qty_commit:.0f} pieces.", ln=True)
    pdf.cell(0, 7, f"The associated financial expenditure for bolts approximates to {convert_to_thousands(overall_cost_blt):.2f} thousand USD.", ln=True)

    pdf.ln(5)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Bolt", "MP17033_BoltQuantity_Graphic_", "1 - Bolt quantity in design versus quantity committed.", x=30, w=150)

    add_section_sub_title("Valve Specifications - SBM Scope")
    pdf.cell(0, 7, f"The cumulative weight of valves stands at approximately {convert_to_tons(total_sbm_valve_weight):.2f} Tons.", ln=True)
    pdf.cell(0, 7, f"In terms of quantity, there are approximately {total_quantity_vlv:.0f} valve units.", ln=True)

    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Valve", "MP17033_ValveWeight_Graphic_", "1 - Valve Weight distribution by the material type.", x=30, w=145)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Valve", "MP17033_ValveQuantity_Graphic_", "2 - Valve total pieces quantity per each material types.", x=30, w=145)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Valve", "MP17033_ValveTotal_Cost_Graphic_", "3 - Valve overall cost per each material types.", x=30, w=145)

    add_section_sub_title("Special Piping Specifications - SBM Scope")
    pdf.cell(0, 7, f"The Special Piping segment boasts a cumulative weight of approximately {convert_to_tons(total_spcpip_sbm_data_weight):.2f} tons.", ln=True)
    pdf.cell(0, 7, f"The quantity of Special Piping units amounts to {total_spcpip_sbm_data_qty:.0f}, with an estimated expenditure of {convert_to_thousands(total_cost_spc):.2f} thousands of dollars.", ln=True)

    pdf.ln(5)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/SPC Piping", "MP17033_SPiping_Cost_Weight_", "1 - Cost and Weight relation of the Special Piping equipments.", x=30, w=150)
    pdf.ln(7)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/SPC Piping", "MP17033_SpecialPiping_Cost_PerSupplier_", "2 - Overall Special Piping cost per different suppliers.", x=30, w=150)

    pdf.add_page()
    pdf.ln(3)
    add_section_title("YARD SCOPE - Breakdown")
    pdf.ln(2)
    add_section_sub_title("Piping Specification - YARD Scope")
    pdf.multi_cell(0, 7, f"The piping under the YARD SCOPE accounts for a total weight of {convert_to_tons(total_piping_yard_net_weight):.2f} tons, "
                         f"comprising {total_qty_commit_pieces_yard:.1f} individual pieces and spanning {total_qty_commit_m_yard:.0f} meters in materials.", align="J")

    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Yard", "MP17033_YARDPiping_Material_weight_", "1 - YARD Piping total weight by different material types.", x=30, w=140)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Yard", "MP17033_YARDPiping_Material_average_", "2 - YARD Piping average weight per material types.", x=30, w=140)

    add_section_sub_title("Special Piping Specifications - YARD Scope")
    pdf.multi_cell(0, 7, f"Within the YARD scope, special piping has a total weight of {convert_to_tons(total_spcpip_yard_data_weight):.2f} tons and consists of {total_spcpip_yard_data_qty:.0f} individual pieces.", align="J")

    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Yard", "MP17033_YARD_SPCPIP_Analyze_", "1 - YARD Special Piping total Pieces and Weight.", x=40, w=121)

    add_section_sub_title("Valve Details - YARD Scope")
    pdf.cell(0, 7, f"For this scope we counted a total Valve Weight of approximately {convert_to_tons(total_yard_valve_weight):.2f} Tons.", ln=True)

    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/PDF Reports/Content", "material_description_breakdown_valve_yard", "1 - YARD Valve weight distribution per material types.", x=45, w=123)

    add_section_sub_title("Bolt Details - YARD Scope")
    pdf.cell(0, 7, f"For the Bolt segment within the YARD scope we have a total of {bolt_yard_data_total_qty_commit:.0f} pieces", ln=True)

    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/PDF Reports/Content", "base_material_breakdown_bolt_yard", "1 - YARD Bolt pieces quantity by material type.", x=50, w=105)

    add_section_sub_title("Structure Specifications - YARD Scope")
    pdf.multi_cell(0, 7, f"The Structural components in the YARD scope amass a total weight of approximately {structure_total_gross_weight:.2f} Tons."
                         f"This equates to about {structure_total_qty_pcs:.0f} individual pieces, spanning an area of {structure_total_qty_m:.2f} meters and covering a surface area of {structure_total_qty_m2:.2f} square meters.", align="J")
    pdf.ln(1)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Structure", "MP17033_Structure_Total Gross Weight_Graphic_", "1 - YARD Structural Gross Weight graphic representation.", x=30, w=145)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Structure", "MP17033_Structure_Wastage Quantity_Graphic_", "2 - Structural Wastage quantity graphic representation.", x=30, w=145)
    add_image_with_legend(pdf, "../Data Pool/DCT Process Results/Graphics/Structure", "MP17033_Structure_Total QTY to commit_Graphic_", "3 - Total YARD Structural quantity committed per material types.", x=30, w=145)

    pdf.add_page()

    # Conclusion
    pdf.ln(5)
    add_section_title("Conclusion")
    pdf.ln(5)
    add_section_sub_title("In-Depth Reflection on Unity Project FPSO Materials Analysis")

    #shorter version of the conclusion
    conclusion_text = ("""
        The thorough examination of materials for the Unity Project FPSO has illuminated critical details about equipment quantity, weight, and associated financial implications. With a total material weight of approximately 3,737.910 metric tons, the project demanded investments in both finances, approximating $1,143,953.07 thousand USD, and manpower, totaling around 620,673.63 hours.

        Distinguishing between the SBM and YARD scopes, we recognized disparities in quantities and weights, each critical in its own right. Piping dominated the weight metrics, spotlighting our engineering focus. Yet, challenges arose: surplus in piping materials, costing an additional $294.931 thousand USD, and wastage in structural equipment, emphasized the need for meticulous planning and resource allocation.

        The project's expansive nature, as seen in the purchase orders placed in NADIA, further showcased our vast procurement operations.

        In essence, the Unity Project FPSO serves as a testament to our commitment, capabilities, and the need for continual optimization. The challenges and successes from this project offer valuable lessons for future endeavors.
        """)

    #Longer version of the conclusion
    text = (f"""The Unity Project FPSO stands as a monumental endeavor for our organization, resonating with profound operational and fiscal connotations. After a comprehensive data analyze, we discerned that the project encapsulated an impressive weight of approximately {total_dct_weight:.3f} metric tons, which sprawls across different equipment ranging from the nuances of small fittings to the grandeur of major structural components.
    Financially, the equipment costs as revealed by the MTO equate to roughly ${total_dct_mto_cost:.3f} thousand USD. However, in its entirety, the Unity Project FPSO attracts a colossal financial commitment of about ${total_project_cost:.2f} thousand USD. This mirrors not just the tangible assets but also our dedication in terms of manpower--reflected by an investment of around {total_project_hours:.2f} hours. The granularity of the project's composition is further illustrated by the sheer number of individual pieces required--{total_dct_pieces:.0f} to be precise, which underscores its intricate nature. The span of materials expands to about {total_dct_meters:.3f} meters, with a significant chunk, approximately {total_dct_meters_yard:.1f} meters, being procured from the YARD.
    
    Diving deeper, a comparative lens between the SBM and YARD scopes showcases disparities in piece count and weight, but both remain important on the project analyze. Under the SBM scope, there are a total of {total_dct_sbm_pieces:.0f} pieces and {total_dct_sbm_weight:.3f} tons, as opposed to the YARD scope with {total_dct_yard_pieces:.0f} pieces and {total_dct_yard_weight:.3f} tons.
    The weight distribution across different material classes elucidates the project's architectural philosophy. With piping taking precedence in weight at {total_dct_piping_weight:.3f} tons, followed by valves at {total_dct_valve_weight:.3f} tons and structural elements at {total_dct_structure_weight:.1f} tons, it's evident where our primary engineering focus lies. The specialized category of piping materials, though niche, plays a pivotal role with its contribution of {total_dct_scp_piping_weight:.3f} tons.
    
    Yet, no endeavor is without its challenges. The presence of a surplus in piping, equivalent to {total_dct_piping_surplus} additional pieces, and the associated cost of ${total_dct_piping_surplus_cost} thousand dollars stresses the essence of refined material management. Moreover, the noted wastage, especially in structural equipment, which amounts to {total_structure_wastage:.1f} tons, calls attention to the imperativeness of optimizing planning processes to curb wastage and uphold fiscal discipline.
    Lastly, a snapshot of the purchase orders placed in NADIA underscores the breadth of our quantity count efforts with {total_dct_po_quantity_piece:.0f} pieces of piping, special piping, and bolt equipment and an expanse of {total_dct_po_quantity_meter:.0f} meters from piping.
    
    In summation, the Unity Project FPSO is not just an engineering marvel but a testament to our capabilities, challenges, and commitment to excellence. As we reflect on its journey, we're reminded of the importance of data-driven insights, strategic material management, and the balance between design intricacy and practical implementation. The learnings derived will indubitably guide our future endeavors.
        """)

    pdf.multi_cell(0, 7, text, align="J")

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Complete MTO Analyze"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"MP17033 DCT Report file.pdf")
    pdf.output(pdf_path)

    return
