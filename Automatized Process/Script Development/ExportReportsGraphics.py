import os
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns


def plot_totals(cost_df, project_number):
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


#Below function not being used
def plot_totals_other(cost_df, project_number):
    commit_totals = cost_df.groupby('Quantity UOM')['Total QTY to commit'].sum()
    po_totals = cost_df.groupby('UOM in PO')['Quantity in PO'].sum()
    currency_totals = cost_df.groupby('Transaction Currency')['PO Cost'].sum()

    # New: calculate cost per weight unit and group by Transaction Currency
    cost_df['Cost Per Weight'] = cost_df['Project Currency Cost'] / (
                cost_df['Total NET weight'] + cost_df['Total Weight using PO quantity'])
    cost_per_weight_totals = cost_df.groupby('Transaction Currency')['Cost Per Weight'].sum()

    # Create graphics directory if not exists
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    fig, axs = plt.subplots(4, 1, figsize=(12, 20))  # Updated: 4 subplots instead of 3

    sns.barplot(x=commit_totals.index, y=commit_totals.values, ax=axs[0])
    axs[0].set_title('Total Committed Quantity per Quantity UOM')
    axs[0].set_xlabel('Quantity UOM')
    axs[0].set_ylabel('Total Committed Quantity')
    # Add value labels
    for p in axs[0].patches:
        axs[0].text(p.get_x() + p.get_width() / 2., p.get_height(),
                    '%d' % int(p.get_height()),
                    fontsize=12, color='black', ha='center', va='bottom')

    sns.barplot(x=po_totals.index, y=po_totals.values, ax=axs[1])
    axs[1].set_title('Total Quantity in PO per UOM in PO')
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

    # New: plot cost per weight unit
    sns.barplot(x=cost_per_weight_totals.index, y=cost_per_weight_totals.values, ax=axs[3])
    axs[3].set_title('Cost per Weight Unit per Transaction Currency')
    axs[3].set_xlabel('Transaction Currency')
    axs[3].set_ylabel('Cost per Weight Unit')

    # Add value labels
    for p in axs[3].patches:
        axs[3].text(p.get_x() + p.get_width() / 2., p.get_height(),
                    '%.2f' % float(p.get_height()),
                    fontsize=12, color='black', ha='center', va='bottom')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_Piping_Analyze_Graphics_Illustration_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


#below function not being used
def plot_cost_per_weight(cost_df, project_number):
    # Calculate cost per weight for each row
    cost_df['Cost per Weight'] = cost_df['Project Currency Cost'] / cost_df['Total Weight using PO quantity']

    # Aggregate totals
    cost_totals = cost_df['Project Currency Cost'].sum()
    net_weight_totals = cost_df['Total NET weight'].sum()
    weight_totals = cost_df['Total Weight using PO quantity'].sum()

    # Calculate total cost per weight
    total_cost_per_weight = cost_totals / weight_totals

    # Create graphics directory if not exists
    graphics_dir = "../Data Pool/DCT Process Results/graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure
    fig, ax = plt.subplots(figsize=(12, 5))

    # Plot total cost per weight
    sns.barplot(x=['Cost per Weight'], y=[total_cost_per_weight], ax=ax)

    ax.set_title('Total Cost per Weight')
    ax.set_ylabel('Cost per Weight')

    # Add value labels
    for p in ax.patches:
        ax.text(p.get_x() + p.get_width() / 2., p.get_height(),
               '%.2f' % float(p.get_height()),
               fontsize=12, color='black', ha='center', va='bottom')

    plt.tight_layout()

    # Save figure
    fig_name = f"MP{project_number}_Total_Cost_Per_Weight.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


def plot_cost_per_weight_and_totals(cost_df, project_number):
    # Calculate cost per weight for each row
    cost_df['Cost per Weight'] = cost_df['Project Currency Cost'] / cost_df['Total Weight using PO quantity']

    # Aggregate totals
    cost_totals = cost_df['Project Currency Cost'].sum()
    net_weight_totals = cost_df['Total NET weight'].sum()
    weight_totals = cost_df['Total Weight using PO quantity'].sum()

    # Calculate total cost per weight
    total_cost_per_weight = cost_totals / weight_totals

    # Create graphics directory if not exists
    graphics_dir = "../Data Pool/DCT Process Results/graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure
    fig, ax = plt.subplots(figsize=(12, 5))

    # Plot total cost per weight
    sns.barplot(x=['Cost per KG', 'Total Cost in Dollars', 'Total Weight in KG'],
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


def plot_material_cost1(cost_df, project_number):
    # Aggregate costs and convert to thousands
    material_costs = cost_df.groupby('Base Material')['Cost'].sum() / 1000

    # Create graphics directory if not exists
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure
    fig, ax = plt.subplots(figsize=(12, 5))

    # Plot total cost per material
    sns.barplot(x=material_costs.index, y=material_costs.values, ax=ax)

    ax.set_title('Total Cost per Material Type')
    ax.set_xlabel('Material Type')
    ax.set_ylabel('Total Cost (Thousands of Dollars)')
    plt.xticks(rotation=60)  # Rotate the x-axis labels to be vertical

    # Add value labels
    for p in ax.patches:
        ax.text(p.get_x() + p.get_width() / 2., p.get_height(),
               '%.2f' % float(p.get_height()),
               fontsize=12, color='black', ha='center', va='bottom')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_PipingTotal_Cost_Material_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


def plot_material_cost(cost_df, project_number):
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
        'SUPER DUPLEX STAINLESS STEEL': 'SDS'
        # Add more mappings as needed...
    }
    # And a dictionary mapping codes back to full names
    code_to_name = {v: k for k, v in name_to_code.items()}

    # Replace the material names in the DataFrame with their codes
    cost_df_copy = cost_df.copy()
    cost_df_copy['Base Material'] = cost_df_copy['Base Material'].map(name_to_code)

    # Aggregate costs and convert to thousands
    material_costs = cost_df_copy.groupby('Base Material')['Cost'].sum() / 1000

    # Create graphics directory if not exists
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

    # Create a legend mapping codes to full names
    handles = [plt.Rectangle((0, 0), 1, 1, color=barplot.patches[i].get_facecolor()) for i in range(len(code_to_name))]
    plt.legend(handles, code_to_name.values(), title='Code to Material Type')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_PipingTotal_Cost_Material_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')
    cost_df['Base Material'] = cost_df['Base Material'].map(code_to_name)  # Add this line




def plot_material_weight(cost_df, project_number):
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
        'SUPER DUPLEX STAINLESS STEEL': 'SDS'
        # Add more mappings as needed...
    }
    # And a dictionary mapping codes back to full names
    code_to_name = {v: k for k, v in name_to_code.items()}

    # Replace the material names in the DataFrame with their codes
    cost_df_copy = cost_df.copy()
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
    cost_df['Base Material'] = cost_df['Base Material'].map(code_to_name)  # Add this line
