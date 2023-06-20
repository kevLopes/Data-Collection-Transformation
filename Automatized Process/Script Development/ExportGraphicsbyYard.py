import matplotlib.pyplot as plt
import seaborn as sns
import os
from datetime import datetime


def plot_material_analyze_piping(df, project_number):
    # Create graphics directory if it doesn't exist
    graphics_dir = "../Data Pool/DCT Process Results/Graphics"
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
        fig_name = f"MP{project_number}_{plot_title.replace(' ', '_')}_{timestamp}.png"
        fig_path = os.path.join(graphics_dir, fig_name)
        plt.savefig(fig_path)
        print(f'Figure saved at {fig_path}')

    # Total QTY to commit per UOM by Material Type
    quantity_df = df.groupby(['Pipe Base Material', 'Quantity UOM'])['Total QTY to commit'].sum().reset_index()
    fig, ax = plt.subplots(figsize=(12, 5))
    sns.barplot(data=quantity_df, x='Pipe Base Material', y='Total QTY to commit', hue='Quantity UOM', ax=ax)
    ax.set_title('Total QTY to Commit per UOM by Material Type')

    # Add value labels on the bars
    for i, bar in enumerate(barplot.patches):
        barplot.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                     f'{bar.get_height():.2f}',
                     ha='center', va='bottom',
                     fontsize=10, color='black')

    plt.tight_layout()

    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    fig_name = f"MP{project_number}_TotalQTY_to_Commit_per_UOM_by_Material_Type_{timestamp}.png"
    fig_path = os.path.join(graphics_dir, fig_name)
    plt.savefig(fig_path)
    print(f'Figure saved at {fig_path}')

    plt.close('all')