from tkinter import filedialog, simpledialog, messagebox
import template_file_handling
from datetime import datetime
import tkinter as tk
import re
import os

sheet_names_file_path = "filename_sheet.txt"

def get_template_file_path():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    # Open file dialog to select a template file
    template_file_path = filedialog.askopenfilename(title="Select Template File",
                                           filetypes=[("Excel files", "*.xlsx")])
    root.destroy()  # Close the root window

    if not template_file_path:
        # If no file is selected, exit the application
        messagebox.showerror("No File Selected", "No file was selected. The application will now exit.")
        exit()
    
    return template_file_path

template_file_path = get_template_file_path()

sheet_file = template_file_handling.load_sheet_names(sheet_names_file_path, template_file_path)

def show_benchmark_warning():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Display the message box
    result = messagebox.askokcancel("Benchmark File Naming",
                                    "Please ensure that each file name starts with the name of the benchmark it represents.\n\nDo you want to proceed?")
    if not result:
        # Exit the application if the user does not confirm
        root.quit()
        root.destroy()
        exit()




# Function to save sheet names to a text file
def save_sheet_names(sheet_names_file_path, sheet_names):
    with open(sheet_names_file_path, 'w') as file:
        file.write('\n'.join(sheet_names))

def update_window_size():
    num_checkboxes = len(checkboxes)
    new_height = 400 + (num_checkboxes * 25)  # Adjust the base height and add extra space for each checkbox
    root.geometry(f"400x{new_height}")
    

def browse_files(sheet_name):
    window_title = f"Choose {sheet_name} Files"
    file_paths = filedialog.askopenfilenames(title=window_title, filetypes=[("CSV files", "*.csv")])
    sorted_file_paths = sorted(list(file_paths), key=lambda x: os.path.basename(x))
    return sorted_file_paths

def get_selected_files():
    selected_files = {}
    for sheet_name, var in checkboxes.items():
        if var.get() == 1:
            files = browse_files(sheet_name)
            if files:
                selected_files[sheet_name] = files
    return selected_files

def browse_directory():
    return filedialog.askdirectory(title="Select destination directory")

def get_template_file_path_string():
    return template_file_path

def update_checkboxes():
    select_all_state = select_all_var.get()
    for var in checkboxes.values():
        var.set(select_all_state)

def submit():
    root.quit()

def add_sheet():
    new_sheet = simpledialog.askstring("Add Sheet", "Enter new sheet name:")
    if new_sheet:
        purchase_date = simpledialog.askstring("Add Sheet", "Enter the purchase date (MM/DD/YYYY):")
        if purchase_date:
            # Validate date format
            if re.match(r"\d{1,2}/\d{1,2}/\d{4}$", purchase_date):
                try:
                    # Parse the date to ensure it's valid
                    datetime.strptime(purchase_date, "%m/%d/%Y")
                    if template_file_handling.add_sheet_to_excel(new_sheet, purchase_date, template_file_path):
                        sheet_file.append(new_sheet)
                        save_sheet_names(sheet_names_file_path, sheet_file)
                        update_checkboxes_list()
                    else:
                        messagebox.showwarning("Unable to Create Sheet", f"Unable to create sheet '{new_sheet}' in Excel, please try closing the template file.")
                except ValueError:
                    messagebox.showwarning("Invalid Date", "Please enter a valid date in MM/DD/YYYY format.")
            else:
                messagebox.showwarning("Invalid Date Format", "Please enter the date in MM/DD/YYYY format.")

def remove_sheet():
    selected_sheets = []
    for sheet_name, var in checkboxes.items():
        if var.get() == 1:
            selected_sheets.append(sheet_name)
    if not selected_sheets:
        messagebox.showinfo("No Sheets Selected", "Please select sheets to remove.")
        return

    confirm_remove = messagebox.askyesno("Confirm Remove", "Are you sure you want to remove the selected sheets?")

    if confirm_remove:
        try:
            for sheet_name in selected_sheets:
                if template_file_handling.remove_sheet_from_excel(sheet_name,template_file_path):
                    sheet_file.remove(sheet_name)
                    del checkboxes[sheet_name]
                else:
                    messagebox.showwarning("Unable to Delete Sheet", f"Unable to delete sheet '{sheet_name}' from Excel, please try closing the template file.")
            save_sheet_names(sheet_names_file_path, sheet_file)
            update_checkboxes_list()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while removing sheets: {e}")


def update_checkboxes_list():
    for widget in checkbox_frame.winfo_children():
        widget.destroy()

    for sheet_name in sheet_file:
        var = tk.IntVar()
        checkboxes[sheet_name] = var
        checkbox = tk.Checkbutton(checkbox_frame, text=sheet_name, variable=var, bg=bg_color, fg=fg_color)
        checkbox.pack(pady=5, anchor="w")
    update_window_size()

def create_selecte_all_checkboxes():
    select_all_checkbox = tk.Checkbutton(root, text="Select All", variable=select_all_var, command=update_checkboxes)
    select_all_checkbox.config(bg="#ffeb3b", fg="#000000", font=("Helvetica", 12, "bold"))
    select_all_checkbox.pack(pady=10, anchor="center")



# Create the main application window
root = tk.Tk()
root.title("File Selection")
root.geometry("400x400")
bg_color = "#f0f0f0"
fg_color = "#333333"
btn_color = "#007bff"
btn_hover_color = "#0056b3"
root.configure(bg=bg_color)

# Frame for checkboxes
checkbox_frame = tk.Frame(root, bg=bg_color)
checkbox_frame.pack(pady=10)

# Checkboxes dictionary
checkboxes = {}
select_all_var = tk.IntVar()

create_selecte_all_checkboxes()
# Update checkboxes with the current list
update_checkboxes_list()

# Add a button to submit selections
submit_button = tk.Button(root, text="Submit", command=submit, bg=btn_color, fg="white", activebackground=btn_hover_color)
submit_button.pack(pady=10)

# Add buttons to add and remove sheets
add_sheet_button = tk.Button(root, text="Add Sheet", command=add_sheet, bg=btn_color, fg="white", activebackground=btn_hover_color)
add_sheet_button.pack(pady=10)

remove_sheet_button = tk.Button(root, text="Remove Sheet", command=remove_sheet, bg=btn_color, fg="white", activebackground=btn_hover_color)
remove_sheet_button.pack(pady=10)

# Start the GUI event loop
root.mainloop()
