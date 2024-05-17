from datetime import datetime
from pathlib import Path
import interface
import openpyxl
import shutil
import csv
import os

template_file_path = "C:\\Users\\Admin\\Desktop\\generates_reports\\template.xlsx"

# Initialize lists
effective_lease_rate_change_TOA = []
effective_revenue_occ_IPRPU = []
occupancy_units_O = []	
propert_value_TAO = []
propert_value_IPRPU = []
propert_value_O = []


def copy_file_to_directory(file_path, destination_directory):
    try:
        # Get the file name from the file path
        file_name = file_path.split('/')[-1]  # Extract the file name from the file path (assuming Unix-like paths)
        
        # Get today's date in YYYY-MM-DD format
        today_date = datetime.now().strftime('%Y-%m-%d')
        
        # Construct the new file name with today's date
        new_file_name = f"generated_report_{today_date}.xlsx"  # Assuming the file extension is ".xlsx"
                        
        # Create the destination file path
        destination_path = f"{destination_directory}/{new_file_name}"
        
        # Check if the destination file already exists
        if os.path.exists(destination_path):
            # Append a suffix to the file name until a unique name is found
            suffix = 1
            while True:
                new_file_name = f"generated_report_{today_date} ({suffix}).xlsx"
                destination_path = os.path.join(destination_directory, new_file_name)
                if not os.path.exists(destination_path):
                    break
                suffix += 1

        
        # Copy the file to the destination directory
        shutil.copy2(file_path, destination_path)
        
        print(f"File '{file_name}' copied to '{destination_path}'.")
        
        return destination_path
    
    except Exception as e:
        print(f"Error occurred: {e}")

def file_exists(file_path):
    return Path(file_path).exists()

def clean_and_convert(values):
    cleaned_values = []
    for value in values:
        try:
            if '%' in value:
                cleaned_value = value.replace('%', '')
                if cleaned_value:
                    float_value = float(cleaned_value) / 100
                    cleaned_values.append(float_value)
            elif '$' in value:
                cleaned_value = value.replace('$', '').replace(',', '')
                if cleaned_value:
                    float_value = float(cleaned_value)
                    cleaned_values.append(float_value)
            else:
                cleaned_value = value.strip()
                if cleaned_value:
                    float_value = float(cleaned_value)
                    cleaned_values.append(float_value)
                else:
                    cleaned_values.append("#N\\A")
                    print(f"Warning: '{value}' is an empty string and will be skipped.")
        except ValueError:
            cleaned_values.append(None)  # Append None instead of an empty string for failed conversions
    return cleaned_values


def find_end_of_needed_data_cell(file_path, target_value):
    with open(file_path, 'r', newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            for column, value in row.items():
                if value == target_value:
                    return [column,reader.line_num]
                    
                
def process_csv(file_path):
    with open(file_path, 'r', newline='') as csvfile:
        reader = csv.reader(csvfile)
        for row_index, row in enumerate(reader, start=1):
            if row_index > 1:
                break 
            for column_index, cell in enumerate(row, start=0):
                if cell == "Trade Out - Effective Lease Rate Change - Comparison":
                    effective_lease_rate_change_TOA.extend([column_index, row_index])
                elif cell == "In Place Rent Per Unit - Effective Revenue (Rev / Occ Unit) - Comparison":
                    effective_revenue_occ_IPRPU.extend([column_index, row_index])
                elif cell == "Occupancy - Occupancy Units - Comparison":
                    occupancy_units_O.extend(([column_index, row_index]))
                elif cell == "In Place Rent Per Unit - Effective Revenue (Rev / Occ Unit) - Value" and not propert_value_IPRPU:
                    propert_value_IPRPU.extend(([column_index, row_index]))
                elif cell == "Trade Out - Effective New Lease Rate Change - Value" and not propert_value_TAO:
                    propert_value_TAO.extend(([column_index, row_index]))
                elif cell == "Occupancy - Occupancy Units - Value" and not propert_value_O:
                    propert_value_O.extend(([column_index, row_index]))
                    


def get_column_data(file_path, column_index, start_row=1, end_row=None):
    column_data = []
    with open(file_path , 'r', newline= '') as csvfile:
        reader = csv.reader(csvfile)
        for row_index, row in enumerate(reader, start =1):
            if row_index < start_row:
                continue
            if end_row is not None and row_index > end_row:
                break
            if len(row) >= column_index:
                column_data.append(row[column_index])
    return column_data            

                

def write_needed_data_to_file(new_file_path, needed_data, sheet_name, start_cell, end_cell):
    # Create a new workbook and select the active sheet
    new_workbook = openpyxl.load_workbook(new_file_path)
    new_sheet = new_workbook[sheet_name]

    start_column, start_row = start_cell
    end_column, end_row = end_cell
    
    if start_row != end_row:
        print("Start and end cells should be in the same row.")
        return

    # Ensure there are enough cells to store the needed data
    if len(needed_data) > (end_column - start_column + 1):

        print("Not enough space to store all needed data.")
        return

    # Write the data from needed_data to the new sheet
    print(clean_and_convert(needed_data))
    for index, value in enumerate(clean_and_convert(needed_data)):
        new_sheet.cell(row=start_row, column=start_column + index, value=value)

    # Save the new workbook
    new_workbook.save(new_file_path)
    print(f"Data written to {new_file_path}")
    

def populate_sheet(file_paths, new_file_path, sheet_name):
    for i in range (len(file_paths)):    
        if not file_exists(file_paths[i]):
            print(f"File not found: {file_paths[i]}")
            return

        process_csv(file_paths[i])
        end_of_needed_data_cell = find_end_of_needed_data_cell(file_paths[i], "SUMMARY")
        
        # Define start and end cell addresses
        start_cell_TAO_initial = [3,3]
        end_cell_TAO_inital = [3,3]

        start_cell_O_initial = [3,9]
        end_cell_O_initial = [3,9]

        start_cell_IPRPU_initial = [3,15]
        end_cell_IPRPU_initial = [3,15]
        
        start_end_cells = [[start_cell_TAO_initial, end_cell_TAO_inital],
                           [start_cell_O_initial, end_cell_O_initial],
                           [start_cell_IPRPU_initial, end_cell_IPRPU_initial]]
        needed_data = []
        # Write the needed data to a new file
        data_columns = [effective_lease_rate_change_TOA, occupancy_units_O,effective_revenue_occ_IPRPU]
        for j in range (len(data_columns)): 
            needed_data.extend([get_column_data(file_paths[i], data_columns[j][0], 2, end_of_needed_data_cell[1]-1)])
        for j in range (len(needed_data)):
            start_end_cells[j][0][1] += i
            start_end_cells[j][1][1] += i
            start_end_cells[j][1][0] += len(needed_data[j])
            write_needed_data_to_file(new_file_path, needed_data[j], sheet_name ,start_end_cells[j][0], start_end_cells[j][1])
        effective_lease_rate_change_TOA.clear()
        effective_revenue_occ_IPRPU.clear()
        occupancy_units_O.clear()

    start_cell_propert_value_TAO_initial =[3,6]
    end_cell_propert_value_TAO_initial =[3,6]
    start_cell_propert_value_O_initial =[3,12]
    end_cell_propert_value_O_initial =[3,12]
    start_cell_propert_value_IPRPU_initial =[3,18]
    end_cell_propert_value_IPRPU_initial =[3,18]
    propert_value_needed_data =[]
    propert_value_data_columns = [propert_value_TAO, propert_value_O,propert_value_IPRPU]

    start_end_cells_propert_value = [[start_cell_propert_value_TAO_initial, end_cell_propert_value_TAO_initial],
                                    [start_cell_propert_value_O_initial,end_cell_propert_value_O_initial],
                                    [start_cell_propert_value_IPRPU_initial,end_cell_propert_value_IPRPU_initial]]
    
    for j in range (len(propert_value_data_columns)): 
        propert_value_needed_data.extend([get_column_data(file_paths[i], propert_value_data_columns[j][0], 2, end_of_needed_data_cell[1]-1)])
        
    for j in range(len(propert_value_needed_data)):
        start_end_cells_propert_value[j][1][0] += len(propert_value_needed_data[j])
        write_needed_data_to_file(new_file_path, propert_value_needed_data[j], sheet_name, start_end_cells_propert_value[j][0], start_end_cells_propert_value[j][1])
        start_end_cells_propert_value[j][0][1] += len(file_paths) + 3
        start_end_cells_propert_value[j][1][1] += len(file_paths) + 3        
    

 
def main():

    # Call the function to get the selected files from the user interface
    selected_files = interface.get_selected_files()

    # Create an empty dictionary to store the selected files
    sheets_with_files = {}

    # Iterate over the selected files dictionary and populate sheets_with_files
    for sheet_name, files in selected_files.items():
        sheets_with_files[sheet_name] = files

    new_file_path = interface.browse_directory()
    
    new_file_path = copy_file_to_directory(template_file_path, new_file_path)
    for sheet_name, file_paths in sheets_with_files.items():
        print("sheet_name", sheet_name)
        populate_sheet(file_paths, new_file_path, sheet_name)

if __name__ == "__main__":
    main()