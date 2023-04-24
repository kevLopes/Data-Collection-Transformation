import tkinter as tk
from tkinter import ttk
import AnalyzeProcessByMaterial
from tkinter import messagebox
from tkinter import filedialog
import EcosysData_API
import DataCollector


class MyWindow:
    def __init__(self, master):
        # Create the main window
        self.master = master
        self.master.title("Data Collection and Transformation")

        # Create a label and a button
        self.label = tk.Label(master, text="Automatic Process", font=("TkDefaultFont", 20), fg="grey")
        self.label.pack(pady=10)

        # Create a frame for the input field and button
        self.frame = tk.Frame(master)
        self.frame.pack(pady=10)

        # Create a label for the input field
        self.input_label = tk.Label(self.frame, text="Project Number:")
        self.input_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # Create an entry field for user input
        self.input_field = tk.Entry(self.frame)
        self.input_field.grid(row=0, column=1, padx=5, pady=5)

        # Create a label for the dropdown list
        self.type_label = tk.Label(self.frame, text="Material Type:")
        self.type_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        # Create a dropdown list for selecting project type
        self.type_var = tk.StringVar(value="Choose Type")
        self.type_dropdown = ttk.OptionMenu(self.frame, self.type_var, "Choose Type", "All Materials", "Piping",
                                            "Valve", "Bolt",
                                            "Special Piping", "Structure", "Bent")
        self.type_dropdown.config(width=16)
        self.type_dropdown.grid(row=1, column=1, padx=5, pady=5)

        # Create a label for the folder path input field
        #self.folder_path_label = tk.Label(self.frame, text="Materials Data Hub Folder:")
        #self.folder_path_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        # Create an entry field for the folder path
        #self.folder_path_input = tk.Entry(self.frame)
        #self.folder_path_input.grid(row=2, column=1, padx=5, pady=5)

        # Create a button to open the folder selection dialog
        #self.folder_path_button = tk.Button(self.frame, text="Browse", command=self.choose_folder_path)
        #self.folder_path_button.grid(row=2, column=2, padx=5, pady=5)

        # Create a check button variable
        self.check_button_var = tk.BooleanVar()

        # Create a label for the check button
        self.check_button_label = tk.Label(self.frame, text="Refresh Ecosys Data")
        self.check_button_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")

        # Create a check button
        self.check_button = tk.Checkbutton(self.frame, variable=self.check_button_var, width=2, height=2,
                                           command=self.on_check_button_clicked)
        self.check_button.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # Create a submit button
        self.submit_button = tk.Button(master, text="Submit Data", command=self.on_submit_clicked)
        self.submit_button.pack(pady=10)

        # Adjust the Submit button layout
        self.adjust_submit_button_layout()

        # Create a clear button
        self.clear_button = tk.Button(master, text="Clear Data", command=self.on_clear_clicked)
        self.clear_button.pack(pady=5)

        # Center the window and widgets
        self.center_window()

    def on_check_button_clicked(self):
        pass

    def choose_folder_path(self):
        # Open the folder selection dialog and get the selected path
        selected_path = filedialog.askdirectory(initialdir="/", title="Folder Path")

        # Update the folder path input field with the selected path
        self.folder_path_input.delete(0, tk.END)
        self.folder_path_input.insert(0, selected_path)

    def adjust_submit_button_layout(self):
        # Change the background color of the Submit button
        self.submit_button.config(bg="lightblue")

        # Change the width of the Submit button
        self.submit_button.config(width=10)

    def center_window(self):
        # Center the window on the screen
        self.master.update_idletasks()
        width = self.master.winfo_width()
        height = self.master.winfo_height()
        x = (self.master.winfo_screenwidth() - width) // 2
        y = (self.master.winfo_screenheight() - height) // 2
        self.master.geometry(f"+{x}+{y}")

        # Center the widgets in the frame
        for widget in self.frame.winfo_children():
            widget.grid_configure(padx=5, pady=5)

    def on_submit_clicked(self):
        # Get the user input and selected project type
        project_number = self.input_field.get()
        material_type = self.type_var.get()
        #folder_path = self.folder_path_input.get()
        folder_path = "../Data Pool/Material Data Organized"

        # Check if all fields have data
        if not project_number or material_type == "Choose Type" or not folder_path:
            # Display a message if fields are not filled in
            tk.messagebox.showerror("Error", "Please fill in the fields")
        else:
            # Check if the check button is checked
            flag = False
            if self.check_button_var.get():
                print("Analyze on going. Data are being refreshed from Ecosys")
                EcosysData_API.ecosys_po_lines_data_api(
                f"https://ecosys-stg.sbmoffshore.com/ecosys/api/restjson/EcosysPOLinesData_DCTAPI_KOL"
                f"/?RootCostObject=MP{project_number}", "keven.deOliveiralope", "My-SBM#code23", project_number)

                #PO Header and SUN transactions API
                #cosysData_API.ecosys_poheader_lines_data_api(
                #f"https://ecosys-stg.sbmoffshore.com/ecosys/api"f"/restjson/EcosysPOHeadersData_DCTAPI_KOL"
                #f"/?RootCostObject=MP{project_number}","keven.deOliveiralope", "My-SBM#code23", project_number)
                #EcosysData_API.ecosys_sun_lines_data_api(
                #f"https://ecosys-stg.sbmoffshore.com/ecosys/api/restjson"f"/EcosysSUNData_DCTAPI_KOL"
                #f"/?RootCostObject=MP{project_number}","keven.deOliveiralope", "My-SBM#code23", project_number)

                #Material Type Piping
                if material_type == "Piping":
                    #Organize Data from Data Hub
                    DataCollector.data_collector_piping(project_number, material_type)
                    # Do something with the input and type (e.g., save to a file or database)
                    distinct_product_codes = AnalyzeProcessByMaterial.extract_distinct_product_codes_piping(folder_path, project_number, material_type)
                    print("Analyze done for all Distinct Material codes:", distinct_product_codes)
                    flag = True
                # Material Type Valve
                elif material_type == "Valve":
                    # Organize Data from Data Hub
                    DataCollector.data_collector_valve(project_number, material_type)
                    # Do something with the input and type (e.g., save to a file or database)
                    distinct_product_codes = AnalyzeProcessByMaterial.extract_distinct_product_codes_valve(folder_path,project_number,material_type)
                    print("Analyze done for all Distinct Material codes:", distinct_product_codes)
                    flag = True
                # Material Type Bolt
                elif material_type == "Bolt":
                    # Organize Data from Data Hub
                    DataCollector.data_collector_bolt(project_number, material_type)
                    # Do something with the input and type (e.g., save to a file or database)
                    flag = True
                # Material Type Structure
                elif material_type == "Structure":
                    print("Implementation not done yet")
                # Material Type Bend
                elif material_type == "Bend":
                    print("Implementation not done yet")
                # All Materials
                elif material_type == "All Materials":
                    print("Implementation not done yet")
            else:
                print("Analyze on going without collecting data from Ecosys")
                #Material Type Piping
                if material_type == "Piping":
                    # Organize Data from Data Hub
                    DataCollector.data_collector_piping(project_number, material_type)
                    distinct_product_codes = AnalyzeProcessByMaterial.extract_distinct_product_codes_piping(folder_path, project_number, material_type)
                    print("Analyze done for all Distinct Material codes:", distinct_product_codes)
                    flag = True
                    # Material Type Valve
                elif material_type == "Valve":
                    # Organize Data from Data Hub
                    DataCollector.data_collector_valve(project_number, material_type)
                    # Do something with the input and type (e.g., save to a file or database)
                    distinct_product_codes = AnalyzeProcessByMaterial.extract_distinct_product_codes_valve(folder_path,project_number,material_type)
                    print("Analyze done for all Distinct Material codes:", distinct_product_codes)
                    flag = True
                # Material Type Bolt
                elif material_type == "Bolt":
                    # Organize Data from Data Hub
                    DataCollector.data_collector_bolt(project_number, material_type)
                    # Do something with the input and type (e.g., save to a file or database)
                    flag = True
                # Material Type Structure
                elif material_type == "Structure":
                    print("Implementation not done yet")
                # Material Type Bend
                elif material_type == "Bend":
                    print("Implementation not done yet")
                #All Materials
                elif material_type == "All Materials":
                    print("Implementation not done yet")

            if flag != True:
                tk.messagebox.showinfo("Process Finished",f"It seems that there was an error with the Process. Reference project number: {project_number}, and the for {material_type} Material")
            else:
                if not EcosysData_API.api_error_track:
                    tk.messagebox.showinfo("Process Finished - Ecosys Data Update",f"Please proceed to DCT Process Results folder to see your result for the project number: {project_number}, and the for {material_type} Material")
                else:
                    tk.messagebox.showinfo("Process Finished - Ecosys Data not Update",f"Please proceed to DCT Process Results folder to see your result for the project number: {project_number}, and the for {material_type} Material. Check error logs for more details")

            # Clear the input field
            self.input_field.delete(0, tk.END)

    def on_clear_clicked(self):
        # Clear the input field
        self.input_field.delete(0, tk.END)
        self.folder_path_input.delete(0, tk.END)

        # Reset the Material Type dropdown
        self.type_var.set("Choose Type")


def screen_menu_main():
    # Create the root window and run the event loop
    root = tk.Tk()

    # Set the window position
    MyWindow(root)

    root.mainloop()
