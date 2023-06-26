import matplotlib.pyplot as plt
import seaborn as sns
import os
from datetime import datetime
import pandas as pd


def plot_material_analyze_piping(df, project_number):
    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics/Yard"
    os.makedirs(graphics_dir, exist_ok=True)

    attributes = {
        'Total NET weight': 'Total Net Weight per Material Type',
        'Unit Weight': 'Total Average Unit Weight per Material Type'
    }

    for attribute, plot_title in attributes.items():
        # Aggregate data
        material_data = df.groupby('Pipe Base Material')[attribute].sum()

        # Create figure
        fig, ax = plt.subplots(figsize=(12, 5))

        # Plot total attribute per material
        barplot = sns.barplot(x=material_data.index, y=material_data.values, ax=ax)

        ax.set_title(plot_title)
        ax.set_xlabel('Material Type')
        ax.set_ylabel(attribute)

        # Add value labels on the bars
        for i, bar in enumerate(barplot.patches):
            barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                         f'{bar.get_height():.2f}',
                         ha='center', va='bottom',
                         fontsize=12, color='black')

        plt.tight_layout()

        # Save figure
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        if plot_title == "Total Average Unit Weight per Material Type":
            fig_name = f"MP{project_number}_YARDPiping_Material_average_{timestamp}.png"
        else:
            fig_name = f"MP{project_number}_YARDPiping_Material_weight_{timestamp}.png"
        fig_path = os.path.join(graphics_dir, fig_name)
        plt.savefig(fig_path)
        print(f'Figure saved at {fig_path}')

    # Total QTY to commit per UOM by Material Type
    quantity_df = df.groupby(['Pipe Base Material', 'Quantity UOM'])['Total QTY to commit'].sum().reset_index()
    fig, ax = plt.subplots(figsize=(12, 5))
    barplot = sns.barplot(data=quantity_df, x='Pipe Base Material', y='Total QTY to commit', hue='Quantity UOM', ax=ax)
    ax.set_title('Total Quantity per UOM by Material Type')

    # Add value labels on the bars
    for i, bar in enumerate(barplot.patches):
        barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                     f'{bar.get_height():.2f}',
                     ha='center', va='bottom',
                     fontsize=10, color='black')

    plt.tight_layout()

    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_YARDPiping_Material_QTY_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')

    plt.close('all')

#                       ------------------------ Valve --------------------


def plot_material_analyze_valves(df, project_number):
    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics/Yard"
    os.makedirs(graphics_dir, exist_ok=True)

    attributes = {
        'Quantity': 'Total Quantity per Material Type',
        'Weight': 'Total Weight per Material Type',
        'SIZE (inch)': 'Total Size per Material Type'
    }

    # Create a dictionary mapping full names to codes
    name_to_code = {
        'CARBON STEEL': 'CST',
        'CHROME-MOLLY': 'CRM',
        'COPPER-NICKEL': 'CUN',
        'DUPLEX STAINLESS STEEL': 'DSS',
        'GALVANIZED CARBON STEEL': 'GST',
        'GRE PIPING': 'GRE',
        'HIGH YIELD CARBON STEEL': 'HST',
        'LOW TEMPERATURE CARBON STEEL': 'LST',
        'METALLIC GASKETS': 'MET',
        'NICKEL ALLOY': 'ICO',
        'NON- METALLIC GASKETS': 'NMT',
        'STAINLESS STEEL': 'SST',
        'SUPER DUPLEX STAINLESS STEEL': 'SDS'
    }

    df['General Material Description'] = df['General Material Description'].map(name_to_code)

    for attribute, plot_title in attributes.items():
        # Aggregate data
        material_data = df.groupby('General Material Description')[attribute].sum()

        # Create figure
        fig, ax = plt.subplots(figsize=(12, 5))

        # Plot total attribute per material
        barplot = sns.barplot(x=material_data.index, y=material_data.values, ax=ax)

        ax.set_title(plot_title)
        ax.set_xlabel('Material Type')
        ax.set_ylabel(attribute)

        # Add value labels on the bars
        for i, bar in enumerate(barplot.patches):
            barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                         f'{bar.get_height():.2f}',
                         ha='center', va='bottom',
                         fontsize=12, color='black')

        # Create a legend mapping material codes to full names
        code_to_name = {v: k for k, v in name_to_code.items()}
        unique_materials = sorted(set(material_data.index))
        handles = [plt.Rectangle((0, 0), 1, 1, color=barplot.patches[i].get_facecolor()) for i in range(len(barplot.patches))]
        plt.legend(handles, [code_to_name[code] for code in unique_materials[:len(barplot.patches)]], title='Material Type Legend')

        plt.tight_layout()

        # Save figure
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        fig_name = f"MP{project_number}_YARD{attribute}_Material_{timestamp}.png"
        plt.savefig(os.path.join(graphics_dir, fig_name))

        # Close the figure
        plt.close('all')

#                       ------------------------ Bolt --------------------


def plot_material_analyze_bolts(df, project_number):
    # Create a dictionary mapping full names to codes
    name_to_code = {
        'CHROME-MOLY': 'CRM',
        'COPPER': 'COP'
    }

    # Replace the 'General Material Description' with code names
    df['General Material Description'] = df['General Material Description'].map(name_to_code)

    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics/Yard"
    os.makedirs(graphics_dir, exist_ok=True)

    # set the plot size
    plt.figure(figsize=(15, 10))

    # Create a bar plot of confirmed quantity per material type
    barplot = sns.barplot(x="General Material Description", y="Quantity Confirmed", data=df)

    # set labels and title
    plt.xlabel('Material Type', fontsize=15)
    plt.ylabel('Total Quantity', fontsize=15)
    plt.title(f'Confirmed Quantity per Material Type', fontsize=15)

    # Show material type codes on x-axis
    unique_materials = df['General Material Description'].unique()
    plt.xticks(range(len(unique_materials)), unique_materials, rotation=0)

    # Create a legend using the short material type codes
    handles = [plt.Rectangle((0, 0), 1, 1, color=barplot.patches[i].get_facecolor()) for i in
               range(len(unique_materials))]
    plt.legend(handles, unique_materials, title='Material Type Legend')

    # Add quantities on the bars
    for i, v in enumerate(df['Quantity Confirmed']):
        plt.text(i, v + 0.5, str(v), color='black', ha='center')

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_YARDBolt_Material_{timestamp}.png"
    plt.savefig(os.path.join(graphics_dir, fig_name))

    print("The plot has been saved successfully.")
    plt.close('all')

#                       ------------------------ Special Piping --------------------


def plot_special_piping_yard(df, project_number):
    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics/Yard"
    os.makedirs(graphics_dir, exist_ok=True)

    # Aggregate data across all descriptions
    total_quantity = df["Qty"].sum()
    total_weight = df["Weight"].sum() / 1000

    # Create dictionary for plotting
    data = {"Category": ["Total Quantity", "Total Weight (TONs)"], "Value": [total_quantity, total_weight]}
    plot_df = pd.DataFrame(data)

    # set the plot size
    plt.figure(figsize=(10, 5))

    # Create a bar plot
    barplot = sns.barplot(x="Category", y="Value", data=plot_df)

    # set labels and title
    plt.xlabel('Category', fontsize=15)
    plt.ylabel('Value', fontsize=15)
    plt.title(f'Special Piping Total Quantity and Weight by YARD', fontsize=15)

    # Add quantities on the bars
    for i, v in enumerate(plot_df["Value"]):
        plt.text(i, v + 0.5, str(round(v, 2)), color='black', ha='center')

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_YARD_SPCPIP_Analyze_{timestamp}.png"
    plt.savefig(os.path.join(graphics_dir, fig_name))

    print("The plot has been saved successfully.")
    plt.close('all')
