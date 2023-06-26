import os
from datetime import date
from fpdf import FPDF
import matplotlib.pyplot as plt


#-------------------------- Piping -----------------------------

#SBM Scope
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

    # Add title
    pdf.cell(0, 10, "MTO Piping Analyze - SBM Scope", ln=True, align="C")

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
    pdf.cell(0, 10, "MTO Piping Analyze - YARD Scope", ln=True, align="C")

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
    pdf_path = os.path.join(directory_path, f"SO{project_number}_MTOPipingReport_YARDscope.pdf")
    pdf.output(pdf_path)