import os
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import CostAnalyzeProcessByMaterial
import numpy as np


def plot_piping_totals(cost_df, project_number):
    commit_totals = cost_df.groupby('Quantity UOM')['Total QTY to commit'].sum()
    po_totals = cost_df.groupby('UOM in PO')['Quantity in PO'].sum()
    currency_totals = cost_df.groupby('Transaction Currency')['PO Cost'].sum() / 1000

    # Create graphics directory if not exists
    graphics_dir = "../Data Pool/DCT Process Results/graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    totals = [(commit_totals, 'Piping Total Committed Quantities per MTO', 'UOM in MTO', 'Committed Quantity'),
              (po_totals, 'Piping Total Quantities per Purchase Order', 'UOM in Purchase Orders', 'Quantity in PO'),
              (currency_totals, 'Piping Total Cost per Currencies', 'Transaction Currency', 'Cost (Thousands of Dollars)')]

    for idx, (total, title, xlabel, ylabel) in enumerate(totals):
        fig, ax = plt.subplots(figsize=(12, 5))
        sns.barplot(x=total.index, y=total.values, ax=ax)
        ax.set_title(title)
        ax.set_xlabel(xlabel)
        ax.set_ylabel(ylabel)

        # Add value labels
        for p in ax.patches:
            ax.text(p.get_x() + p.get_width() / 2., p.get_height(),
                   '%d' % int(p.get_height()),
                   fontsize=12, color='black', ha='center', va='bottom')

        plt.tight_layout()

        # Save figure
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        fig_name = f"MP{project_number}_Piping_Graphics_{idx}_{timestamp}.png"
        fig_path = os.path.join(graphics_dir, fig_name)
        plt.savefig(fig_path)
        print(f'Figure saved at {fig_path}')

    plt.close('all') # Close all figures at the end


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
    sns.barplot(x=['Cost per KG', 'Total Cost (Thousands of Dollars)', 'Total Weight in TON'],
                y=[total_cost_per_weight, cost_totals, weight_totals], ax=ax)

    ax.set_title('Piping Representation for Weight and Cost')
    ax.set_ylabel('Values')

    # Add value labels
    for p in ax.patches:
        ax.text(p.get_x() + p.get_width() / 2., p.get_height(),
               '%.2f' % float(p.get_height()),
               fontsize=12, color='black', ha='center', va='bottom')

    plt.tight_layout()
    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_Piping_Cost_Weight_Graphic_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


def plot_piping_cost_per_po(cost_df, project_number):
    folder_path = "../Data Pool/Ecosys API Data/PO Headers"
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"

    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]

    matching_files = [f for f in excel_files if str(project_number) in f]

    if matching_files:
        most_recent_poheader = CostAnalyzeProcessByMaterial.get_most_recent_file(folder_path, matching_files)
        file_path = os.path.join(folder_path, most_recent_poheader)
        df = pd.read_excel(file_path)

        cost_df['Cost'] = cost_df['Cost'] / 1000  # Convert to thousands

        # Remove NaN values, strip whitespace, and filter out empty strings
        cost_df['PO Number'] = cost_df['PO Number'].dropna().str.strip()
        cost_df = cost_df[cost_df['PO Number'] != '']

        # Get total costs by PO Number
        po_totals = cost_df.groupby('PO Number')['Cost'].sum()

        # Create a DataFrame to hold PO Numbers, Costs, and Supplier Names
        final_df = pd.DataFrame(columns=['PO Number', 'Cost', 'Supplier Name'])

        # Iterate over PO totals
        for po_number, cost in po_totals.items():
            # Convert the PO Number to the format in the PO Header file
            po_number_converted = po_number.split('||')[1].replace('-', '.')

            # Find corresponding Supplier Name in df
            supplier_series = df.loc[df['PO Number'] == po_number_converted, 'Supplier Name']

            if not supplier_series.empty:
                supplier_name = supplier_series.iloc[0]

                # Append to final_df
                final_df = final_df.append({
                    'PO Number': po_number,
                    'Cost': cost,
                    'Supplier Name': supplier_name
                }, ignore_index=True)
            else:
                # Handle the case when there is no matching Supplier Name
                # For example, you could add a row with 'Unknown' as Supplier Name
                final_df = final_df.append({
                    'PO Number': po_number,
                    'Cost': cost,
                    'Supplier Name': 'Unknown'
                }, ignore_index=True)

        os.makedirs(graphics_dir, exist_ok=True)

        # Plot total cost by Supplier Name
        supplier_totals = final_df.groupby('Supplier Name')['Cost'].sum().sort_values(ascending=False)
        fig, axs = plt.subplots(figsize=(12, 8))
        sns.barplot(y=supplier_totals.index, x=supplier_totals.values, ax=axs, orient='h')
        axs.set_title('Piping Total Cost per Supplier')
        axs.set_xlabel('Total Cost (Thousands of Dollars)')
        axs.set_ylabel('Supplier Name')

        # Add cost total to the end of each bar
        for p in axs.patches:
            axs.annotate(format(p.get_width(), '.2f'),
                            (p.get_width(), p.get_y() + p.get_height() / 2.),
                            ha='center',
                            va='center',
                            size=10,
                            xytext=(20, 0),
                            textcoords='offset points')

        plt.tight_layout()

        # Save figure
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        fig_name = f"MP{project_number}_Piping_CostPerSupplier_Graphic_{timestamp}.png"
        fig_path = os.path.join(graphics_dir, fig_name)
        plt.savefig(fig_path)
        plt.close() # Close the current plot
        print(f'Figure saved at {fig_path}')

        # Plot total cost by PO Number
        fig, axs = plt.subplots(figsize=(12, 8))
        sns.barplot(y=final_df['PO Number'], x=final_df['Cost'], ax=axs, orient='h')
        axs.set_title('Piping Total Cost per PO Number')
        axs.set_xlabel('Total Cost (Thousands of Dollars)')
        axs.set_ylabel('PO Number')

        # Add cost total to the end of each bar
        for p in axs.patches:
            axs.annotate(format(p.get_width(), '.2f'),
                            (p.get_width(), p.get_y() + p.get_height() / 2.),
                            ha='center',
                            va='center',
                            size=10,
                            xytext=(20, 0),
                            textcoords='offset points')

        plt.tight_layout()

        # Save figure
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        fig_name = f"MP{project_number}_Piping_CostPerPO_Graphic_{timestamp}.png"
        fig_path = os.path.join(graphics_dir, fig_name)
        plt.savefig(fig_path)
        plt.close() # Close the current plot
        print(f'Figure saved at {fig_path}')



def plot_piping_material_cost(cost_df_mt, project_number):
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

    ax.set_title('Piping Total Cost per Material Type')
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

    ax.set_title('Piping Total Net Weight per Material Type')
    ax.set_xlabel('Material Type Code')
    ax.set_ylabel('Total Net Weight (TONs)')

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

#-------------------- Special Piping ---------------


#all images separated
def plot_special_piping_cost_per_po_and_supplier(cost_df, project_number):
    folder_path = "../Data Pool/Ecosys API Data/PO Headers"
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"

    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]
    matching_files = [f for f in excel_files if str(project_number) in f]

    # Convert cost to thousands for better readability
    cost_df['Cost'] = cost_df['Cost'] / 1000

    metrics = ['Cost', 'Quantity in PO', 'MTO Weight']  # Metrics to be plotted

    # Create graphics directory if it doesn't exist
    os.makedirs(graphics_dir, exist_ok=True)

    # Create a DataFrame to hold PO Numbers, Costs, Quantity, Weight and Supplier Names
    final_df = pd.DataFrame(columns=['PO Number', 'Cost', 'Quantity in PO', 'MTO Weight', 'Supplier Name'])

    if matching_files:
        most_recent_poheader = CostAnalyzeProcessByMaterial.get_most_recent_file(folder_path, matching_files)
        file_path = os.path.join(folder_path, most_recent_poheader)
        df = pd.read_excel(file_path)

        # Iterate over each metric
        for metric in metrics:
            po_totals = cost_df.groupby('PO Number')[metric].sum()

            for po_number, metric_value in po_totals.items():
                # Convert the PO Number to the format in the PO Header file
                po_number_converted = po_number.split('||')[1].replace('-', '.')

                # Find corresponding Supplier Name in df
                supplier_series = df.loc[df['PO Number'] == po_number_converted, 'Supplier Name']

                if not supplier_series.empty:
                    supplier_name = supplier_series.iloc[0]

                    if po_number not in final_df['PO Number'].values:
                        final_df = final_df.append({
                            'PO Number': po_number,
                            metric: metric_value,
                            'Supplier Name': supplier_name
                        }, ignore_index=True)
                    else:
                        final_df.loc[final_df['PO Number'] == po_number, metric] = metric_value
                else:
                    # Handle the case when there is no matching Supplier Name
                    if po_number not in final_df['PO Number'].values:
                        final_df = final_df.append({
                            'PO Number': po_number,
                            metric: metric_value,
                            'Supplier Name': 'Unknown'
                        }, ignore_index=True)
                    else:
                        final_df.loc[final_df['PO Number'] == po_number, metric] = metric_value

            # Create figure for current metric per PO Number
            fig, ax = plt.subplots(figsize=(12, 8))
            # Sort for better visibility
            data_to_plot = final_df.sort_values(by=metric, ascending=False)

            # Plot current metric per PO Number
            sns.barplot(y=data_to_plot['PO Number'], x=data_to_plot[metric], ax=ax, orient='h')
            ax.set_title(f'Special Piping Total {metric} per PO Number')
            ax.set_xlabel(f'Total {metric}')
            ax.set_ylabel('PO Number')

            # Add value labels
            for p in ax.patches:
                ax.text(p.get_width(), p.get_y() + p.get_height() / 2.,
                        '%.2f' % float(p.get_width()),
                        fontsize=12, color='black', va='center')

            plt.tight_layout()

            # Save the figure
            timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
            fig_name = f"MP{project_number}_SpecialPiping_{metric}_PerPO_{timestamp}.png"
            fig_path = os.path.join(graphics_dir, fig_name)
            plt.savefig(fig_path)
            print(f'Figure saved at {fig_path}')

            plt.close(fig)

            # Create figure for current metric per Supplier
            fig, ax = plt.subplots(figsize=(12, 8))

            # Aggregate and sort for better visibility
            supplier_totals = final_df.groupby('Supplier Name')[metric].sum()
            data_to_plot = supplier_totals.sort_values(ascending=False).reset_index()

            # Plot current metric per Supplier
            sns.barplot(y=data_to_plot['Supplier Name'], x=data_to_plot[metric], ax=ax, orient='h')
            ax.set_title(f'Special Piping Total {metric} per Supplier')
            ax.set_xlabel(f'Total {metric}')
            ax.set_ylabel('Supplier Name')

            # Add value labels
            for p in ax.patches:
                ax.text(p.get_width(), p.get_y() + p.get_height() / 2.,
                        '%.2f' % float(p.get_width()),
                        fontsize=12, color='black', va='center')

            plt.tight_layout()

            # Save the figure
            timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
            fig_name = f"MP{project_number}_SpecialPiping_{metric}_PerSupplier_{timestamp}.png"
            fig_path = os.path.join(graphics_dir, fig_name)
            plt.savefig(fig_path)
            print(f'Figure saved at {fig_path}')

            plt.close(fig)


#Analyse with 1 pictures for all POs and supplier
def plot_special_piping_cost_per_po1(cost_df, project_number):
    folder_path = "../Data Pool/Ecosys API Data/PO Headers"
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"

    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") or f.endswith(".xls")]
    matching_files = [f for f in excel_files if str(project_number) in f]

    # Convert cost to thousands for better readability
    cost_df['Cost'] = cost_df['Cost'] / 1000

    metrics = ['Cost', 'Quantity in PO', 'MTO Weight']  # Metrics to be plotted

    # Create graphics directory if it doesn't exist
    os.makedirs(graphics_dir, exist_ok=True)

    # Create a DataFrame to hold PO Numbers, Costs, Quantity, Weight and Supplier Names
    final_df = pd.DataFrame(columns=['PO Number', 'Cost', 'Quantity in PO', 'MTO Weight', 'Supplier Name'])

    if matching_files:
        most_recent_poheader = CostAnalyzeProcessByMaterial.get_most_recent_file(folder_path, matching_files)
        file_path = os.path.join(folder_path, most_recent_poheader)
        df = pd.read_excel(file_path)

        # Iterate over each metric
        for metric in metrics:
            po_totals = cost_df.groupby('PO Number')[metric].sum()

            for po_number, metric_value in po_totals.items():
                # Convert the PO Number to the format in the PO Header file
                po_number_converted = po_number.split('||')[1].replace('-', '.')

                # Find corresponding Supplier Name in df
                supplier_series = df.loc[df['PO Number'] == po_number_converted, 'Supplier Name']

                if not supplier_series.empty:
                    supplier_name = supplier_series.iloc[0]

                    if po_number not in final_df['PO Number'].values:
                        final_df = final_df.append({
                            'PO Number': po_number,
                            metric: metric_value,
                            'Supplier Name': supplier_name
                        }, ignore_index=True)
                    else:
                        final_df.loc[final_df['PO Number'] == po_number, metric] = metric_value
                else:
                    # Handle the case when there is no matching Supplier Name
                    if po_number not in final_df['PO Number'].values:
                        final_df = final_df.append({
                            'PO Number': po_number,
                            metric: metric_value,
                            'Supplier Name': 'Unknown'
                        }, ignore_index=True)
                    else:
                        final_df.loc[final_df['PO Number'] == po_number, metric] = metric_value

    # Create a single figure with three subplots (arranged vertically)
    fig, axes = plt.subplots(nrows=3, figsize=(12, 24))

    # Iterate over each metric for plotting
    for i, metric in enumerate(metrics):
        # Sort for better visibility
        data_to_plot = final_df.sort_values(by=metric, ascending=False)

        # Plot on the ith subplot
        ax = axes[i]
        sns.barplot(y=data_to_plot['PO Number'], x=data_to_plot[metric], ax=ax, orient='h')
        ax.set_title(f'Special Piping Total {metric} per PO Number')
        ax.set_xlabel(f'Total {metric}')
        ax.set_ylabel('PO Number')

        # Add value labels
        for p in ax.patches:
            ax.text(p.get_width(), p.get_y() + p.get_height() / 2.,
                    '%.2f' % float(p.get_width()),
                    fontsize=12, color='black', va='center')

    plt.tight_layout()

    # Save the figure with all three subplots
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_SpecialPiping_Metrics_PerPO_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')

    plt.close(fig)

    # Create a separate figure for total cost by supplier
    fig, ax = plt.subplots(figsize=(12, 8))
    supplier_totals = final_df.groupby('Supplier Name')['Cost'].sum().sort_values(ascending=False)
    sns.barplot(y=supplier_totals.index, x=supplier_totals.values, ax=ax, orient='h')
    ax.set_title('Special Piping Total Cost per Supplier')
    ax.set_xlabel('Total Cost (Thousands of Dollars)')
    ax.set_ylabel('Supplier Name')

    # Add value labels
    for p in ax.patches:
        ax.text(p.get_width(), p.get_y() + p.get_height() / 2.,
                '%.2f' % float(p.get_width()),
                fontsize=12, color='black', va='center')

    plt.tight_layout()

    # Save this separate figure
    fig_name = f"MP{project_number}_SpecialPiping_Cost_PerSupplier_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')


#3 separated image
def plot_special_piping_cost_per_po1(cost_df, project_number):
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"

    # Convert cost to thousands for better readability
    cost_df['Cost'] = cost_df['Cost'] / 1000

    metrics = ['Cost', 'Quantity in Purchase Order', 'MTO Weight']  # Metrics to be plotted

    # Create graphics directory if it doesn't exist
    os.makedirs(graphics_dir, exist_ok=True)

    # Iterate over each metric
    for metric in metrics:
        # Group by PO Number and calculate the sum of the current metric
        po_totals = cost_df.groupby('PO Number')[metric].sum().sort_values(ascending=False)

        # Create the plot
        fig, ax = plt.subplots(figsize=(12, 8))
        sns.barplot(y=po_totals.index, x=po_totals.values, ax=ax, orient='h')
        ax.set_title(f'Special Piping Total {metric} per PO Number')
        ax.set_xlabel(f'Total {metric} ')#(if Cost then in Thousands of Dollars)')
        ax.set_ylabel('PO Number')

        # Add value labels
        for p in ax.patches:
            ax.text(p.get_width(), p.get_y() + p.get_height() / 2.,
                    '%.2f' % float(p.get_width()),
                    fontsize=12, color='black', va='center')
        plt.tight_layout()

        # Save the figure
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        fig_name = f"MP{project_number}_SpecialPiping_{metric}_PerPO_{timestamp}.png"
        fig_path = os.path.join(graphics_dir, fig_name)
        plt.savefig(fig_path)
        print(f'Figure saved at {fig_path}')


def plot_special_piping_cost_per_weight_and_totals(cost_df, project_number):
    # Calculate cost per weight for each row
    cost_df['Cost per Weight'] = cost_df['Project Currency Cost'] / cost_df['Total Weight using PO quantity']

    # Aggregate totals
    cost_totals = cost_df['Project Currency Cost'].sum() / 1000
    weight_totals = cost_df['Total Weight using PO quantity'].sum() / 1000

    # Calculate total cost per weight
    total_cost_per_weight = (cost_totals / weight_totals) if weight_totals else np.nan

    # Calculate total cost per currency
    total_cost_per_currency = cost_df.groupby('Transaction Currency')['Project Currency Cost'].sum().div(1000)

    # Create graphics directory if not exists
    graphics_dir = "../Data Pool/DCT Process Results/graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Create figure for total cost per weight
    fig, axs = plt.subplots(figsize=(12, 6))
    sns.barplot(x=['Cost per KG', 'Total Cost (Thousands of Dollars)', 'Total Weight in TON'],
                y=[total_cost_per_weight, cost_totals, weight_totals], ax=axs)
    axs.set_title('Special Piping Cost per Weight')
    axs.set_ylabel('Values')

    # Add value labels
    for p in axs.patches:
        axs.text(p.get_x() + p.get_width() / 2., p.get_height(),
                 '%.2f' % float(p.get_height()),
                 fontsize=12, color='black', ha='center', va='bottom')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_SPiping_Cost_Weight_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    plt.close()  # Close the current plot
    print(f'Figure saved at {fig_path}')

    # Create figure for total cost per currency
    fig, axs = plt.subplots(figsize=(12, 6))
    sns.barplot(x=total_cost_per_currency.index, y=total_cost_per_currency.values, ax=axs)
    axs.set_title('Special Piping Cost per Currency')
    axs.set_ylabel('Cost (Thousands of Dollars)')

    # Add value labels
    for p in axs.patches:
        axs.text(p.get_x() + p.get_width() / 2., p.get_height(),
                 '%.2f' % float(p.get_height()),
                 fontsize=12, color='black', ha='center', va='bottom')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_SPiping_Cost_Currency_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    plt.close()  # Close the current plot
    print(f'Figure saved at {fig_path}')

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

    ax.set_title('Valve Total Cost per Material Type')
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
    fig_name = f"MP{project_number}_ValveTotal_Cost_Graphic_{timestamp}.png"
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

    # Plot total quantity per material
    fig, ax1 = plt.subplots(figsize=(12, 6))
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

    # Create a legend for the plot
    code_to_name = {v: k for k, v in name_to_code.items()}
    unique_materials = sorted(name_to_code.values())
    handles = [plt.Rectangle((0, 0), 1, 1, color=barplot1.patches[i].get_facecolor()) for i in
               range(len(unique_materials))]
    ax1.legend(handles, [code_to_name[code] for code in unique_materials], title='Material Type Legend')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_ValveQuantity_Graphic_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    plt.close()  # Close the current plot
    print(f'Figure saved at {fig_path}')

    # Plot total weight per material
    fig, ax2 = plt.subplots(figsize=(12, 6))
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

    # Create a legend for the plot
    code_to_name = {v: k for k, v in name_to_code.items()}
    unique_materials = sorted(name_to_code.values())
    handles = [plt.Rectangle((0, 0), 1, 1, color=barplot2.patches[i].get_facecolor()) for i in
               range(len(unique_materials))]
    ax2.legend(handles, [code_to_name[code] for code in unique_materials], title='Material Type Legend')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_ValveWeight_Graphic_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    plt.close()  # Close the current plot
    print(f'Figure saved at {fig_path}')


def plot_valve_cost_quantity_per_currency(cost_df_currency, project_number):
    # Aggregate cost and quantity per currency
    currency_cost = cost_df_currency.groupby('Transaction Currency')['Cost'].sum() / 1000  # in thousands
    currency_quantity = cost_df_currency.groupby('Transaction Currency')['Quantity'].sum()

    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"
    os.makedirs(graphics_dir, exist_ok=True)

    # Plot total cost per currency
    fig, ax1 = plt.subplots(figsize=(12, 6))
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

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_Valve_Cost_Graphic_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    plt.close()  # Close the current plot
    print(f'Figure saved at {fig_path}')

    # Plot total quantity per currency
    fig, ax2 = plt.subplots(figsize=(12, 6))
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
    fig_name = f"MP{project_number}_Valve_Quantity_Graphic_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    plt.close()  # Close the current plot
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

    ax.set_title('Bolt Total Cost per Material Type')
    ax.set_xlabel('Material Type')
    ax.set_ylabel('Total Cost')

    # Add value labels on the bars
    for i, bar in enumerate(barplot.patches):
        barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                     f'{bar.get_height():.2f}',
                     ha='center', va='bottom',
                     fontsize=12, color='black')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_BoltTotal_Cost_Graphic_{timestamp}.png"
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
    plt.title('Bolt difference between "Quantity Commit" and "Quantity Confirmed"')

    # Add quantity labels on the bars
    for i, bar in enumerate(barplot.patches):
        barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                     f'{bar.get_height():.2f}',
                     ha='center', va='bottom',
                     fontsize=12, color='black')

    plt.tight_layout()

    # Save figure
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_BoltQuantity_Graphic_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')
