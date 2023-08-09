import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import Submit_button_actions


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
        self.type_label = tk.Label(self.frame, text="Equipment Type:")
        self.type_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        # Create a dropdown list for selecting project type
        self.type_var = tk.StringVar(value="Choose Type")
        self.type_dropdown = ttk.OptionMenu(self.frame, self.type_var, "Choose Type", "All Equipment", "Piping",
                                            "Valve", "Bolt",
                                            "Special Piping", "Structure", "Bend")
        self.type_dropdown.config(width=16)
        self.type_dropdown.grid(row=1, column=1, padx=5, pady=5)

        # Create a check button variable for Refresh Ecosys Data
        self.check_button_var = tk.BooleanVar()

        # Create a label for the Refresh Ecosys Data check button
        self.check_button_label = tk.Label(self.frame, text="Refresh Ecosys Data")
        self.check_button_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")

        # Create a check button for Refresh Ecosys Data
        self.check_button = tk.Checkbutton(self.frame, variable=self.check_button_var, width=2, height=2,
                                           command=self.on_check_button_clicked)
        self.check_button.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # Create a check button variable for Include YARD Scope
        self.yard_scope_var = tk.BooleanVar()

        # Create a label for the Include YARD Scope check button
        self.yard_scope_label = tk.Label(self.frame, text="Include YARD Scope")
        self.yard_scope_label.grid(row=4, column=0, padx=5, pady=5, sticky="w")

        # Create a check button for Include YARD Scope
        self.yard_scope_button = tk.Checkbutton(self.frame, variable=self.yard_scope_var, width=2, height=2,
                                                command=self.on_yard_scope_clicked)
        self.yard_scope_button.grid(row=4, column=1, padx=5, pady=5, sticky="w")

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

    def on_yard_scope_clicked(self):
        # This function will run when the "Include YARD Scope" check button is clicked.
        # Implement the functionality here
        pass


    def on_check_button_clicked(self):
        pass

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

        # Check if all fields have data
        if not project_number or material_type == "Choose Type":
            # Display a message if fields are not filled in
            tk.messagebox.showerror("Error", "Please fill in the fields")
        else:
            # Check if the check button is checked
            api_flag = False
            if self.check_button_var.get():
                Submit_button_actions.action_for_ecosys_api(project_number)
                api_flag = True

            flag = Submit_button_actions.action_for_material_analyze(project_number, material_type)

            if self.yard_scope_var.get() and material_type != "All Equipment":
                Submit_button_actions.action_for_material_analyze_by_yard(project_number, material_type)

            if not flag:
                tk.messagebox.showinfo("Process Finished", f"It seems that there was an error with the Process. Reference project number: {project_number}, and the for {material_type} Material")
            else:
                if api_flag:
                    tk.messagebox.showinfo("Process Finished - Ecosys Data Update", f"Please proceed to DCT Process Results folder to see your result for the project number: {project_number}, and the for {material_type} Material")
                else:
                    tk.messagebox.showinfo("Process Finished - Ecosys Data not Update", f"Please proceed to DCT Process Results folder to see your result for the project number: {project_number}, and the for {material_type} Material. Check error logs for more details")

            # Clear the input field
            self.input_field.delete(0, tk.END)

    def on_clear_clicked(self):
        # Clear the input field
        self.input_field.delete(0, tk.END)
        #self.folder_path_input.delete(0, tk.END)

        # Reset the Material Type dropdown
        self.type_var.set("Choose Type")


def screen_menu_main():
    # Create the root window and run the event loop
    root = tk.Tk()

    # Set the window position
    MyWindow(root)

    root.mainloop()
