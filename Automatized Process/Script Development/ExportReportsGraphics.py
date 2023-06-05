import os
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd


def plot_piping_totals(cost_df, project_number):
    commit_totals = cost_df.groupby('Quantity UOM')['Total QTY to commit'].sum()
    po_totals = cost_df.groupby('UOM in PO')['Quantity in PO'].sum()
    currency_totals = cost_df.groupby('Transaction Currency')['PO Cost'].sum()

    # Create graphics directory if not exists
    graphics_dir = "../Data Pool/DCT Process Results/graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    fig, axs = plt.subplots(3, 1, figsize=(12, 15))

    sns.barplot(x=commit_totals.index, y=commit_totals.values, ax=axs[0])
    axs[0].set_title('Total Committed Quantity per MTO')
    axs[0].set_xlabel('Quantity UOM')
    axs[0].set_ylabel('Total Committed Quantity')

    # Add value labels
    for p in axs[0].patches:
        axs[0].text(p.get_x() + p.get_width() / 2., p.get_height(),
               '%d' % int(p.get_height()),
               fontsize=12, color='black', ha='center', va='bottom')

    sns.barplot(x=po_totals.index, y=po_totals.values, ax=axs[1])
    axs[1].set_title('Total Quantity from Purchase Order')
    axs[1].set_xlabel('UOM in PO')
    axs[1].set_ylabel('Total Quantity in PO')

    # Add value labels
    for p in axs[1].patches:
        axs[1].text(p.get_x() + p.get_width() / 2., p.get_height(),
               '%d' % int(p.get_height()),
               fontsize=12, color='black', ha='center', va='bottom')

    sns.barplot(x=currency_totals.index, y=currency_totals.values, ax=axs[2])
    axs[2].set_title('Total Cost per Transaction Currency')
    axs[2].set_xlabel('Transaction Currency')
    axs[2].set_ylabel('Total Cost')

    # Add value labels
    for p in axs[2].patches:
        axs[2].text(p.get_x() + p.get_width() / 2., p.get_height(),
               '%d' % int(p.get_height()),
               fontsize=12, color='black', ha='center', va='bottom')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_PipingAnalyze_Graphics_Illustration_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


def plot_piping_cost_per_weight_and_totals(cost_df, project_number):
    # Calculate cost per weight for each row
    cost_df['Cost per Weight'] = cost_df['Project Currency Cost'] / cost_df['Total Weight using PO quantity']

    # Aggregate totals
    cost_totals = cost_df['Project Currency Cost'].sum() / 1000
    net_weight_totals = cost_df['Total NET weight'].sum()
    weight_totals = cost_df['Total Weight using PO quantity'].sum() / 1000

    # Calculate total cost per weight
    total_cost_per_weight = (cost_totals / weight_totals)

    # Create graphics directory if not exists
    graphics_dir = "../Data Pool/DCT Process Results/graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure
    fig, ax = plt.subplots(figsize=(12, 5))

    # Plot total cost per weight
    sns.barplot(x=['Cost per KG', 'Total Cost in Thousands of USD', 'Total Weight in TON'],
                y=[total_cost_per_weight, cost_totals, weight_totals], ax=ax)

    ax.set_title('Graphic representation for Weight and Cost')
    ax.set_ylabel('Values')

    # Add value labels
    for p in ax.patches:
        ax.text(p.get_x() + p.get_width() / 2., p.get_height(),
               '%.2f' % float(p.get_height()),
               fontsize=12, color='black', ha='center', va='bottom')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_Piping_Cost_and_Weight_Illustration_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


def plot_piping_material_cost(cost_df_mt, project_number):
    # Create a dictionary mapping full names to codes
    name_to_code = {
        'CARBON STEEL': 'CST',
        'CHROME-MOLY': 'CRM',
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
        # Add more mappings as needed...
    }

    # Replace the material names in the DataFrame with their codes
    cost_df_copy = cost_df_mt.copy()
    cost_df_copy['Base Material'] = cost_df_copy['Base Material'].map(name_to_code)

    # Sort the DataFrame by the material codes
    cost_df_copy.sort_values(by='Base Material', inplace=True)

    # Aggregate costs and convert to thousands
    material_costs = cost_df_copy.groupby('Base Material')['Cost'].sum() / 1000

    # Sort the material costs based on the material codes
    material_costs = material_costs.reindex(sorted(name_to_code.values()))

    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure
    fig, ax = plt.subplots(figsize=(12, 5))

    # Plot total cost per material
    barplot = sns.barplot(x=material_costs.index, y=material_costs.values, ax=ax)

    ax.set_title('Total Cost per Material Type')
    ax.set_xlabel('Material Type Code')
    ax.set_ylabel('Total Cost (Thousands of Dollars)')

    # Add value labels on the bars
    for i, bar in enumerate(barplot.patches):
        barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                     f'{bar.get_height():.2f}',
                     ha='center', va='bottom',
                     fontsize=12, color='black')

    # Create a legend mapping material codes to full names
    code_to_name = {v: k for k, v in name_to_code.items()}
    unique_materials = sorted(name_to_code.values())
    handles = [plt.Rectangle((0, 0), 1, 1, color=barplot.patches[i].get_facecolor()) for i in range(len(unique_materials))]
    plt.legend(handles, [code_to_name[code] for code in unique_materials], title='Material Type Legend')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_PipingTotal_Cost_Material_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')



def plot_piping_material_weight(cost_df_mw, project_number):
    # Create a dictionary mapping full names to codes
    name_to_code = {
        'CARBON STEEL': 'CST',
        'CHROME-MOLY': 'CRM',
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
        # Add more mappings as needed...
    }
    # And a dictionary mapping codes back to full names
    code_to_name = {v: k for k, v in name_to_code.items()}

    # Replace the material names in the DataFrame with their codes
    cost_df_copy = cost_df_mw.copy()
    cost_df_copy['Base Material'] = cost_df_copy['Base Material'].map(name_to_code)

    # Aggregate weights and convert to tons
    material_weights = cost_df_copy.drop_duplicates(subset='Base Material', keep='first').set_index('Base Material')[
                           'Total NET weight'] / 1000

    # Create graphics directory if not exists
    graphics_dir = "../Data Pool/DCT Process Results/graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure
    fig, ax = plt.subplots(figsize=(12, 5))

    # Plot total weight per material
    barplot = sns.barplot(x=material_weights.index, y=material_weights.values, ax=ax)

    ax.set_title('Total Net Weight per Material Type')
    ax.set_xlabel('Material Type Code')
    ax.set_ylabel('Total Net Weight (TONs)')
    #plt.xticks(rotation=75)  # Rotate the x-axis labels to be vertical

    # Add value labels on the bars
    for i, bar in enumerate(barplot.patches):
        barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                     f'{bar.get_height():.2f}',
                     ha='center', va='bottom',
                     fontsize=12, color='black')

    # Create a legend mapping codes to full names
    handles = [plt.Rectangle((0, 0), 1, 1, color=barplot.patches[i].get_facecolor()) for i in range(len(code_to_name))]
    plt.legend(handles, code_to_name.values(), title='Material Type Legend')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_PipingTotal_Net_Weight_Per_Material_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')
    #cost_df_copy['Base Material'] = cost_df_copy['Base Material'].map(code_to_name)  # Add this line


#-------------------- Valve --------------------
def plot_valve_material_cost(cost_df_mt, project_number):
    # Create a dictionary mapping full names to codes
    name_to_code = {
        'CARBON STEEL': 'CST',
        'DUCTILE IRON': 'DIA',
        'COPPER': 'COP',
        'DUPLEX STAINLESS STEEL': 'DSS',
        'Ni-Aluminum Bronze': 'ALB',
        'Nodular CI': 'NCI',
        'HIGH YIELD CARBON STEEL': 'HST',
        'LOW TEMPERATURE CARBON STEEL': 'LST',
        'BRONZE': 'BRO',
        'NICKEL ALLOY': 'ICO',
        'ALUMINUM BRONZE': 'ABR',
        'STAINLESS STEEL': 'SST',
        'SUPER DUPLEX STAINLESS STEEL': 'SDS'
        # Add more mappings as needed...
    }

    # Replace the material names in the DataFrame with their codes
    cost_df_copy = cost_df_mt.copy()
    cost_df_copy['Base Material'] = cost_df_copy['Base Material'].map(name_to_code)

    # Aggregate costs and convert to thousands
    material_costs = cost_df_copy.groupby('Base Material')['Cost'].sum() / 1000

    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure
    fig, ax = plt.subplots(figsize=(12, 5))

    # Plot total cost per material
    barplot = sns.barplot(x=material_costs.index, y=material_costs.values, ax=ax)

    ax.set_title('Total Cost per Material Type')
    ax.set_xlabel('Material Type Code')
    ax.set_ylabel('Total Cost (Thousands of Dollars)')

    # Add value labels on the bars
    for i, bar in enumerate(barplot.patches):
        barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                     f'{bar.get_height():.2f}',
                     ha='center', va='bottom',
                     fontsize=12, color='black')

    # Create a legend mapping material codes to full names
    code_to_name = {v: k for k, v in name_to_code.items()}
    unique_materials = sorted(name_to_code.values())
    handles = [plt.Rectangle((0, 0), 1, 1, color=barplot.patches[i].get_facecolor()) for i in range(len(unique_materials))]
    plt.legend(handles, [code_to_name[code] for code in unique_materials], title='Material Type Legend')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_ValveTotal_Cost_Illustration_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


def plot_valve_material_quantity_weight(cost_df_mt, project_number):
    # Create a dictionary mapping full names to codes
    name_to_code = {
        'CARBON STEEL': 'CST',
        'DUCTILE IRON': 'DIA',
        'COPPER': 'COP',
        'DUPLEX STAINLESS STEEL': 'DSS',
        'Ni-Aluminum Bronze': 'ALB',
        'Nodular CI': 'NCI',
        'HIGH YIELD CARBON STEEL': 'HST',
        'LOW TEMPERATURE CARBON STEEL': 'LST',
        'BRONZE': 'BRO',
        'NICKEL ALLOY': 'ICO',
        'ALUMINUM BRONZE': 'ABR',
        'STAINLESS STEEL': 'SST',
        'SUPER DUPLEX STAINLESS STEEL': 'SDS'
        # Add more mappings as needed...
    }

    # Replace the material names in the DataFrame with their codes
    cost_df_copy = cost_df_mt.copy()
    cost_df_copy['Base Material'] = cost_df_copy['Base Material'].map(name_to_code)

    # Aggregate quantity and weight
    material_quantity = cost_df_copy.groupby('Base Material')['Quantity'].sum()
    material_weight = cost_df_copy.groupby('Base Material')['Weight'].sum() / 1000

    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure with two subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))

    # Plot total quantity per material
    barplot1 = sns.barplot(x=material_quantity.index, y=material_quantity.values, ax=ax1)

    ax1.set_title('Total Valve Quantity per Material Type')
    ax1.set_xlabel('Material Type Code')
    ax1.set_ylabel('Total Quantity')

    # Add value labels on the bars
    for i, bar in enumerate(barplot1.patches):
        barplot1.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                      f'{bar.get_height():.2f}',
                      ha='center', va='bottom',
                      fontsize=12, color='black')

    # Create a legend for the first subplot
    code_to_name = {v: k for k, v in name_to_code.items()}
    unique_materials = sorted(name_to_code.values())
    handles = [plt.Rectangle((0, 0), 1, 1, color=barplot1.patches[i].get_facecolor()) for i in
               range(len(unique_materials))]
    ax1.legend(handles, [code_to_name[code] for code in unique_materials], title='Material Type Legend')

    # Plot total weight per material
    barplot2 = sns.barplot(x=material_weight.index, y=material_weight.values, ax=ax2)

    ax2.set_title('Total Valve Weight per Material Type')
    ax2.set_xlabel('Material Type Code')
    ax2.set_ylabel('Total Weight in TONs')

    # Add value labels on the bars
    for i, bar in enumerate(barplot2.patches):
        barplot2.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                      f'{bar.get_height():.2f}',
                      ha='center', va='bottom',
                      fontsize=12, color='black')

    # Create a legend for the second subplot
    code_to_name = {v: k for k, v in name_to_code.items()}
    unique_materials = sorted(name_to_code.values())
    handles = [plt.Rectangle((0, 0), 1, 1, color=barplot2.patches[i].get_facecolor()) for i in
               range(len(unique_materials))]
    ax2.legend(handles, [code_to_name[code] for code in unique_materials], title='Material Type Legend')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_ValveQuantityWeight_Illustration_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


def plot_valve_cost_quantity_per_currency(cost_df_currency, project_number):
    # Aggregate cost and quantity per currency
    currency_cost = cost_df_currency.groupby('Transaction Currency')['Cost'].sum() / 1000  # in thousands
    currency_quantity = cost_df_currency.groupby('Transaction Currency')['Quantity'].sum()

    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure with two subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))

    # Plot total cost per currency
    barplot1 = sns.barplot(x=currency_cost.index, y=currency_cost.values, ax=ax1)
    ax1.set_title('Total Valve Cost per Currency')
    ax1.set_xlabel('Currency')
    ax1.set_ylabel('Total Cost (Thousands of Dollars)')

    # Add value labels on the bars
    for i, bar in enumerate(barplot1.patches):
        barplot1.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                      f'{bar.get_height():.2f}',
                      ha='center', va='bottom',
                      fontsize=12, color='black')

    # Plot total quantity per currency
    barplot2 = sns.barplot(x=currency_quantity.index, y=currency_quantity.values, ax=ax2)
    ax2.set_title('Total Valve Quantity per Currency')
    ax2.set_xlabel('Currency')
    ax2.set_ylabel('Total Quantity')

    # Add value labels on the bars
    for i, bar in enumerate(barplot2.patches):
        barplot2.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                      f'{bar.get_height():.2f}',
                      ha='center', va='bottom',
                      fontsize=12, color='black')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_Valve_CostQuantity_Currency_Illustration_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


#-------------------- Bolt --------------------
def plot_bolt_material_cost(cost_df_mt, project_number):

    # Replace the material names in the DataFrame with their codes
    cost_df_copy = cost_df_mt.copy()
    cost_df_copy['Base Material'] = cost_df_copy['Base Material']

    # Aggregate costs and convert to thousands
    material_costs = cost_df_copy.groupby('Base Material')['Cost'].sum() / 1000

    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure
    fig, ax = plt.subplots(figsize=(12, 5))

    # Plot total cost per material
    barplot = sns.barplot(x=material_costs.index, y=material_costs.values, ax=ax)

    ax.set_title('Total Cost per Material Type')
    ax.set_xlabel('Material Type')
    ax.set_ylabel('Total Cost (Thousands of Dollars)')

    # Add value labels on the bars
    for i, bar in enumerate(barplot.patches):
        barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                     f'{bar.get_height():.2f}',
                     ha='center', va='bottom',
                     fontsize=12, color='black')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_BoltTotal_Cost_Illustration_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


def plot_bolt_quantity_difference(material_info, project_number):
    # Create a DataFrame from the material info dictionary
    data = []
    for material, info in material_info.items():
        data.append({
            'Material': material,
            'Total QTY to commit': info['Total QTY to commit'],
            'Qty confirmed in design': info['Qty confirmed in design']
        })
    df = pd.DataFrame(data)

    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Melt the DataFrame to make it suitable for seaborn
    df_melted = df.melt(id_vars='Material', var_name='Quantity Type', value_name='Quantity')

    # Create a barplot
    plt.figure(figsize=(12, 8))
    barplot = sns.barplot(x='Material', y='Quantity', hue='Quantity Type', data=df_melted)

    # Add labels and title
    plt.xlabel('Material Type')
    plt.ylabel('Quantity')
    plt.title('Difference between "Quantity Commit" and "Quantity Confirmed"')

    # Add quantity labels on the bars
    for i, bar in enumerate(barplot.patches):
        barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                     f'{bar.get_height():.2f}',
                     ha='center', va='bottom',
                     fontsize=12, color='black')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_BoltQuantity_Illustration_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')
