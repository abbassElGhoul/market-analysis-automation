import tkinter as tk
from tkinter import filedialog


def browse_files(sheet_name):
    window_title = f"Choose {sheet_name} Files"
    file_paths = filedialog.askopenfilenames(title = window_title)
    # Convert file_paths tuple to a list
    file_paths_list = list(file_paths)
    return file_paths_list

def get_selected_files():
    selected_files = {}
    for sheet_name, var in checkboxes.items():
        if var.get() == 1:  # Checkbox is checked
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
        
def check_select_all():
    all_selected = all(var.get() == 1 for var in checkboxes.values())
    select_all_var.set(1 if all_selected else 0)
    

def submit():
    root.quit()  # Exit the GUI event loop

# Create the main application window
root = tk.Tk()
root.title("File Selection")
root.geometry("400x300")  # Set initial window size

# Use a consistent color scheme
bg_color = "#f0f0f0"  # Light gray background
fg_color = "#333333"  # Dark gray foreground
btn_color = "#007bff"  # Blue button color
btn_hover_color = "#0056b3"  # Darker blue hover color

root.configure(bg=bg_color)

# Checkbox labels
sheet_names = [
    "Coventry",
    "Tribeca",
    "FHF",
    "Aspen",
    "Columbia",
    "Concord"
]

# Checkboxes dictionary to store the checkbox variables
checkboxes = {}

# Add a "Select All" checkbox
select_all_var = tk.IntVar()
select_all_checkbox = tk.Checkbutton(root, text="Select All", variable=select_all_var, command=update_checkboxes, bg=bg_color, fg=fg_color)
select_all_checkbox.config(font=("Helvetica", 12, "bold"))
select_all_checkbox.pack(pady=10, anchor="center")


# Add checkboxes for each sheet name
for sheet_name in sheet_names:
    var = tk.IntVar()
    checkboxes[sheet_name] = var
    checkbox = tk.Checkbutton(root, text=sheet_name, variable=var, bg=bg_color, fg=fg_color)
    checkbox.pack(pady=5, anchor="w")

# Add a button to submit selections
submit_button = tk.Button(root, text="Submit", command=submit, bg=btn_color, fg="white", activebackground=btn_hover_color)
submit_button.pack(pady=10)



root.mainloop()  