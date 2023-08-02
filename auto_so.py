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
    Creates and populates the "raw data" sheet with data from the service files.
    Population of data is by row/vessel entry.

    Parameters:
    output_file (str): The name of the output file.
    service_files (list): List of service file names; datasource
    raw_cells_to_extract (dict): Map of column headers to extract with cell references of the data.
    input_dir (str): The name of the directory containing the input files


    """
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
    sheet_name = "raw data"
    raw_data_sheet = create_sheet(workbook, sheet_name)
    workbook.active = raw_data_sheet
    workbook.title = "Service Overview"
    raw_data_sheet.append(raw_headers_list)

    # Iterate through the .xlsx files and extract cell data
    for service_file in service_files:
        file_path = os.path.join(input_dir, service_file)
        row_data = {key: None for key in raw_headers_list}
            
        # SERVICE DESC
        lookup = "SERVICE DESC"
        cell_reference = raw_cells_to_extract[lookup]
        row_data[lookup] = extract_cell(file_path, cell_reference)

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

        # Create and populate "raw data" sheet
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

