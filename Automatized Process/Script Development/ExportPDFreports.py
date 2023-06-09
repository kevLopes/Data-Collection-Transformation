import os
from datetime import date
from fpdf import FPDF
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns


#-------------------------- Piping -----------------------------

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
            pdf.cell(0, 10, f"Total QTY to commit (METER): {filtered_data_meter['Total QTY to commit'].sum()}", ln=True, align="L")
            total_qty_mtrs += filtered_data_meter['Total QTY to commit'].sum()

        if not filtered_data_pcs.empty:
            pdf.cell(0, 10, f"Total QTY to commit (PCE): {filtered_data_pcs['Total QTY to commit'].sum()}", ln=True, align="L")
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
    pdf.cell(0, 10, "Carbon steel is the most prevalent material type, accounting for a significant portion of the total records.", ln=True, align="L")
    pdf.cell(0, 10, "Chrome-moly and nickel alloy materials have relatively low representation in the dataset.", ln=True, align="L")
    pdf.cell(0, 10, "The total net weight varies across different materials, with carbon steel and duplex stainless steel having the highest values.", ln=True, align="L")
    pdf.cell(0, 10, "The average unit weight also shows variation, with chrome-moly and non-metallic gaskets having higher average weights.", ln=True, align="L")
    pdf.cell(0, 10, "The quantity committed differs based on the material type and is reported in both meters and pieces.", ln=True, align="L")

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Piping"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"{project_number}_MTOPipingReport_SBMscope.pdf")
    pdf.output(pdf_path)


#SBM Scope
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


#YARD Scope
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
            pdf.cell(0, 10, f"Total QTY to commit (METER): {filtered_data_meter['Total QTY to commit'].sum()}", ln=True, align="L")
            total_qty_mtrs += filtered_data_meter['Total QTY to commit'].sum()

        if not filtered_data_pcs.empty:
            pdf.cell(0, 10, f"Total QTY to commit (PCE): {filtered_data_pcs['Total QTY to commit'].sum()}", ln=True, align="L")
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


#-------------------------- Valve -----------------------------

#SBM Scope
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


#YARD Scope
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


#------------------------------------------- Bolt --------------------------------------------------

#SBM Scope
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


#YARD Scope
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


#----------------------------------- STRUCTURE -------------------------------------------


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
    pdf.cell(0, 10, f"Overall Total Gross Weight: {filtered_df['Total Gross Weight'].sum():.4f} (TONs)", ln=True, align="L")
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
        pdf.cell(0, 10, f"Total Quantity including Wastage: {group['Quantity Including Wastage'].sum():.4f}", ln=True, align="L")
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
    pdf.cell(0, 10, "The analysis of equipment data for Unity Project provides the following insights:", ln=True, align="L")
    pdf.cell(0, 10, f"- Primary Steel, Plate is the most common material type, accounting for {primary_plate_percentage:.4f} % of the records.", ln=True, align="L")
    pdf.cell(0, 10, f"- Secondary Steel, Profiles have a relatively high representation, accounting for {sec_plate_percentage:.4f} % of the records.", ln=True, align="L")
    pdf.cell(0, 10, f"- Special Steel, Tubular have the lowest representation, accounting for {spc_plate_percentage:.4f} % of the records.", ln=True, align="L")
    pdf.cell(0, 10, f"- The average unit weight across all materials is {filtered_df['Unit Weight'].mean():.4f} kg.", ln=True, align="L")
    pdf.cell(0, 10, f"- The overall total gross weight is {filtered_df['Total Gross Weight'].sum():.4f} tons.", ln=True, align="L")
    pdf.cell(0, 10, f"- The total quantity committed is {filtered_df['Total QTY to commit'].sum():.4f}.", ln=True, align="L")
    pdf.cell(0, 10, "- The data breakdown by material type and the corresponding quantities provide a detailed overview of the equipment distribution.", ln=True, align="L")

    # Create the directory if it doesn't exist
    directory_path = "../Data Pool/DCT Process Results/PDF Reports/Structure"
    os.makedirs(directory_path, exist_ok=True)

    # Save the PDF report
    pdf_path = os.path.join(directory_path, f"{project_number}_MTOStructureReport_SBMscope.pdf")
    pdf.output(pdf_path)