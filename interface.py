import tkinter as tk
import template_file_handling
from tkinter import filedialog, simpledialog, messagebox

sheet_names_file_path = "filename_sheet.txt"

# Function to load sheet names from a text file


sheet_file = template_file_handling.load_sheet_names(sheet_names_file_path)

# Function to save sheet names to a text file
def save_sheet_names(file_path, sheet_names):
    with open(file_path, 'w') as file:
        file.write('\n'.join(sheet_names))

def update_window_size():
    num_checkboxes = len(checkboxes)
    new_height = 400 + (num_checkboxes * 25)  # Adjust the base height and add extra space for each checkbox
    root.geometry(f"400x{new_height}")
    

def browse_files(sheet_name):
    window_title = f"Choose {sheet_name} Files"
    file_paths = filedialog.askopenfilenames(title=window_title)
    file_paths_list = list(file_paths)
    return file_paths_list

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

def update_checkboxes():
    select_all_state = select_all_var.get()
    for var in checkboxes.values():
        var.set(select_all_state)

def submit():
    root.quit()

def add_sheet():
    new_sheet = simpledialog.askstring("Add Sheet", "Enter new sheet name:")
    if new_sheet:
        if template_file_handling.add_sheet_to_excel(new_sheet):    
            sheet_file.append(new_sheet)
            save_sheet_names(sheet_names_file_path, sheet_file)
            update_checkboxes_list()
        else:
            messagebox.showwarning("Unable to Create Sheet", f"Unable to Create sheet '{new_sheet}' in Excel, please try closing the template file.")


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
                if template_file_handling.remove_sheet_from_excel(sheet_name):
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
