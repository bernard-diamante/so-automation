import os
import shutil

import pandas as pd
from openpyxl import Workbook, load_workbook

def create_directory(dir_name):
    """
    Creates a directory in the same parent directory with a given name.

    Parameters:
    dir_name (str): The name of the directory to create.
                    Default is "xlsx" for the converted Excel files.
    """
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
    else:
        shutil.rmtree(dir_name)
        os.makedirs(dir_name)

def create_sheet(output_workbook, sheet_name):
    """
    Creates a sheet in the Excel workbook with a given name

    Parameters:
    outbook_workbook (Workbook): The workbook where the sheet will be created.
    sheet_name (str): The name of the sheet to be created.

    Returns:
    Worksheet: Created sheet
    """
    output_sheet = output_workbook.create_sheet(sheet_name)
    return output_sheet

def convert_xls_to_xlsx(input_folder, output_folder):
    """
    Converts all .xls files in a directory, outputs
    .xlsx to another directory, and returns a list
    of the output file names.

    Parameters:
    input_folder (str): File path of directory containing .xls files.
    output_folder (str): File path of directory where .xlsx files are created.

    Returns:
    list: The list of .xlsx file names created

    Raises:
    ValueError: If there are no service files found in the directory.
    """
    xls_files = os.listdir(input_folder)
    xls_files = [file for file in xls_files
                 if file.startswith("Service_")
                 and file.endswith(".xls")]
    if not xls_files:
        raise ValueError(f"No service files found in the following directory: {input_folder}")


    for xls_file in xls_files:
        input_path = os.path.join(input_folder, xls_file)
        output_path = os.path.join(output_folder, xls_file.replace(".xls", ".xlsx"))

        # Copy over data into the new .xlsx file
        df_xls = pd.read_excel(input_path)

        wb_xlsx = Workbook()
        sheet_xlsx = wb_xlsx.active

        for row in df_xls.iterrows():
            sheet_xlsx.append(row[1].tolist())

        wb_xlsx.save(output_path)

    service_files = [file for file in os.listdir(output_folder)]

    return service_files

def extract_cell(file_path, cell_reference):
    """
    Returns data from a cell in the given file found
    through the input file path.

    Parameters:
    file_path (str): Relative path to the .xlsx input file.
    cell_reference (str): The location of the cell to extract.

    Returns:
    Various: The data from the cell extracted.
    """
    try:
        workbook = load_workbook(file_path)
        sheet = workbook.active
        if cell_reference == None:
            return None
        return sheet[cell_reference].value
    except Exception as e:
        print(f"Error: {e}")

def populate_raw_data_sheet(output_file, service_files, input_dir):
    """
    Creates and populates the "raw" sheet with data from the service files.
    Population of data is by row/vessel entry.

    Parameters:
    output_file (str): The name of the output file.
    service_files (list): List of service file names; datasource
    raw_cells_to_extract (dict): Map of column headers to extract with cell references of the data.
    input_dir (str): The name of the directory containing the input files


    """

    def find_cell_value_to_right(file_path, search_string):
        """
        Finds the specified search string and returns the value to the right.
        """
        try:
            workbook = load_workbook(file_path)
            sheet = workbook.active

            for row in sheet.iter_rows(values_only=True):
                if row[2] == search_string:  # Column C is index 2 (0-based index)
                    # Return value from column D (index 3)
                    return row[3]
            
            # If the search_string is not found, return None
            return None
        except Exception as e:
            print(f"Error: {e}")
            return None
    
    def find_cell_value_to_below(file_path, search_string):
        workbook = load_workbook(file_path)
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=30, min_col=3):
            for cell in row:
                if cell.value == search_string:
                    row_index = cell.row
                    column_index = cell.column
                    value_below = sheet.cell(row=row_index+1, column=column_index).value
                    return value_below

        return None  # Keyword not found or cell below is empty

    def list_vesselnames_cell_references(file_path):
        """
        Given an Excel file, return a list of all 
        cell references of the vessel names.

        Parameters:
        file_path (str): Relative path to the .xlsx input file.
        """
        workbook = load_workbook(file_path)
        sheet = workbook.active
        start_extraction = False
        extracted_cells = []

        for row_num, row in enumerate(sheet.iter_rows(min_row=1, min_col=3, max_col=3, values_only=True), start=1):
            cell_value = row[0]
            if start_extraction:
                if cell_value is None:
                    break
                extracted_cells.append(f"C{row_num}")
            elif cell_value == "Vessel name":
                start_extraction = True

        return extracted_cells

    # Field list
    raw_cells_to_extract = {
        "SERVICE NAME": None,
        "SERVICE DESC": "D3",
        "ROUTE": None,
        "LEAD SL": None,
        "SAILING FREQ": None,
        "PARTICIPANTS": None,
        "VESSEL OPERATOR":None,
        "# OF VESSELS": None,
        "# OF VESSELS PER ROW COUNT": None,
        "WEEKLY CAPACITY":None,
        "SHIPS USED": None,
        "PORT ROTATION": None,
        "VESSEL SIZE": None,
        "VESSEL NAME": None
        }
    raw_headers_internal = [
        "PORT",
        "MICT SERVICE NAME",
        "ALT SRVC CD"
        ]
    
    raw_headers_extract = [cell for cell in raw_cells_to_extract.keys()]
    raw_headers_list = list(raw_headers_internal[:2]) + list(raw_headers_extract[:12]) + [raw_headers_internal[2]] + list(raw_headers_extract[12:])
    # Create sheet and populate header
    workbook = load_workbook(output_file)
    sheet_name = "raw"
    raw_data_sheet = create_sheet(workbook, sheet_name)
    workbook.active = raw_data_sheet
    workbook.title = "Service Overview"
    raw_data_sheet.append(raw_headers_list)
    raw_data_sheet.freeze_panes = "A2"

    # Iterate through the .xlsx files and extract cell data
    for service_file in service_files:
        file_path = os.path.join(input_dir, service_file)
        row_data = {key: None for key in raw_headers_list}
            
        # SERVICE DESC
        lookup = "SERVICE DESC"
        cell_reference = raw_cells_to_extract[lookup]
        cell_value = extract_cell(file_path, cell_reference)
        row_data[lookup] = cell_value

        # SERVICE NAME
        def has_parentheses(s):
            return '(' in s and ')' in s
        
        def has_dash(s):
            return '-' in s or "–" in s or "—" in s
        
        def get_text_in_last_parentheses(text):
            last_open = text.rfind("(")
            last_close = text.rfind(")")
            if last_open != -1 and last_close > last_open:
                return text[last_open + 1:last_close]
            else:
                return ""
        
        def get_text_after_last_dash(text):
            last_dash_idx = 0
            idx_temp = 0
            for char in ["-", "–", "—"]:
                temp = service_desc.rfind(char)
                if temp > idx_temp:
                    last_dash_idx = temp
            extracted_text = service_desc[last_dash_idx+1:].strip()
            return extracted_text

        # SERVICE NAME
        lookup = "SERVICE NAME"
        service_desc = row_data["SERVICE DESC"]
        if has_parentheses(service_desc):
            cell_value = get_text_in_last_parentheses(service_desc)
            if has_dash(cell_value):
                cell_value = get_text_after_last_dash(cell_value)[:-1]
        else:
            cell_value = get_text_after_last_dash(service_desc)
        row_data[lookup] = cell_value
        
        # ROUTE
        lookup = "ROUTE"
        cell_value = find_cell_value_to_right(file_path, "Coverage")
        row_data[lookup] = cell_value

        # SAILING FREQ
        lookup = "SAILING FREQ"
        cell_value = find_cell_value_to_right(file_path, "Sailing frequency")
        row_data[lookup] = cell_value


        # WEEKLY CAPACITY
        lookup = "WEEKLY CAPACITY"
        cell_value = find_cell_value_to_right(file_path, "Weekly capacity (teu)")
        row_data[lookup] = cell_value

        # SHIPS USED
        lookup = "SHIPS USED"
        cell_value = find_cell_value_to_right(file_path, "Proforma fleet")
        row_data[lookup] = cell_value

        def extract_vessel_size(input_string):
            # Split the input_string at the word "from"
            parts = input_string.split("from", 1)

            # Check if "from" exists in the input_string
            if len(parts) > 1:
                # Extract everything after "from" and remove leading/trailing spaces
                result = parts[1].strip()
                return result
            else:
                # If "from" is not found, return None
                return None
            
        # VESSEL SIZE
        lookup = "VESSEL SIZE"
        try:
            cell_value = extract_vessel_size(row_data["SHIPS USED"])[:-1]
        except TypeError:
            cell_value = "-"
        row_data[lookup] = cell_value

        # # OF VESSELS
        lookup = "# OF VESSELS"
        try:
            cell_value = int(row_data["SHIPS USED"].split()[0])
        except ValueError   :
            cell_value = row_data["SHIPS USED"].split()[0]

        row_data[lookup] = cell_value

        # # OF VESSELS PER ROW COUNT
        lookup = "# OF VESSELS PER ROW COUNT"
        cell_value = 1
        row_data[lookup] = cell_value

        # PORT ROTATION
        lookup = "PORT ROTATION"
        cell_value = find_cell_value_to_below(file_path, "Port rotation")
        row_data[lookup] = cell_value

        # WEEKLY CAPACITY
        lookup = "WEEKLY CAPACITY"
        cell_value = find_cell_value_to_right(file_path, "Weekly capacity (teu)")
        row_data[lookup] = cell_value


        # VESSEL NAME, VESSEL OPERATOR
        lookup = "VESSEL NAME"
        operator = "VESSEL OPERATOR"
        vessel_name_coordinates_list = list_vesselnames_cell_references(file_path)
        # If no vessels listed, default to -
        if not vessel_name_coordinates_list:
            cell_value = "-"
            row_data[lookup] = cell_value
            raw_data_sheet.append([value for value in row_data.values()])
        else:
            # Fill unique fields (VESSEL NAME, VESSEL OPERATOR)
            for cell_reference in vessel_name_coordinates_list:
                cell_value = extract_cell(file_path, cell_reference)
                row_data[lookup] = cell_value

                cell_reference = "K" + cell_reference[1:]
                cell_value = extract_cell(file_path, cell_reference)
                row_data[operator] = cell_value
            
                raw_data_sheet.append([value for value in row_data.values()])

        


    workbook.save(output_file)
    return workbook

def main():
    try:
        xls_dir = "xls"
        xlsx_dir = "xlsx"
        output_file = "Service Overview.xlsx"

        # Convert downloaded .xls to .xlsx Service files
        create_directory("xlsx")
        service_files = convert_xls_to_xlsx(xls_dir, xlsx_dir)
        
        # Create Service Overview.xlsx output file
        output_workbook = Workbook()
        output_workbook.save(output_file)

        # Create and populate "raw" sheet
        output_workbook = populate_raw_data_sheet(output_file, service_files, xlsx_dir)
        
        # Remove the sheet created by default
        default_sheet = output_workbook["Sheet"]
        output_workbook.remove(default_sheet)
        output_workbook.save(output_file)

        
    # Handle exceptions
    except FileNotFoundError as fnf_error:
        print(f"Error: {fnf_error}. Please enter a valid username")

    except Exception as e:
        print(f"An error occured: {e}")


if __name__ == "__main__":
    main()

