import os
from openpyxl import Workbook, load_workbook
import pandas as pd
import shutil

def create_directory(converted_dir = "xlsx"):
    if not os.path.exists(converted_dir):
        os.makedirs(converted_dir)
    else:
        shutil.rmtree(converted_dir)
        os.makedirs(converted_dir)

def convert_xls_to_xlsx(input_folder, output_folder):
    """
    Converts all .xls files in a directory, outputs
    .xlsx to another directory, and returns the result.

    Parameters:
    input_folder (str): File path of directory containing .xls files.
    output_folder (str): File path of directory where .xlsx files are created.

    Returns:
    list: The list of .xlsx file names created

    Raises:
    ValueError: If there are no service files found in the directory.
    """
    xls_files = os.listdir(input_folder)
    xls_files = [file for file in xls_files if file.startswith("Service_") and file.endswith(".xls")]
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

def extract_cell(file_path, cell):
    """
    Returns data from a cell in the given file.

    Parameters:
    file_path (str): Path to the .xlsx file
    cell (str): The name of the cell to extract.

    Returns:
    Various: The data from the cell extracted.
    """
    try:
        if cell == None:
            return None
        workbook = load_workbook(file_path)
        sheet = workbook.active
        return sheet[cell].value
    except Exception as e:
        print(f"Error: {e}")

def main():
    try:
        username = str(input("Please input your username: "))
        # root_path = f"C:\\Users\\{username}\\Desktop\\"
        input_folder = "xls"
        output_folder = "xlsx"
        output_file = "Service Overview.xlsx"
        output_sheet = "raw data"

        # Field list
        cells_to_extract = {
            "SERVICE NAME": None,
            "SERVICE DESC": None,
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
            "VESSEL_NAME": None
            }
        headers_internal = [
            "PORT",
            "MICT SERVICE NAME",
            "ALT SRVC CD"
            ]
        
        headers_extract = [cell for cell in cells_to_extract.keys()]
        cells_extract = [cell for cell in cells_to_extract.values()]
        headers_list = list(headers_internal[:2]) + list(headers_extract[:12]) + [headers_internal[2]] + list(headers_extract[12:])

        # .xlsx Service files
        create_directory()
        service_files = convert_xls_to_xlsx(input_folder, output_folder)
        
        # Create Service Overview.xlsx output file and fill with headers
        output_workbook = Workbook()
        output_workbook.remove(output_workbook.active)
        output_sheet = output_workbook.create_sheet(output_sheet)
        output_sheet.append(headers_list)

        for service_file in service_files:
            file_path = os.path.join(output_folder, service_file)
            output_sheet.append([extract_cell(file_path, cell) for cell in cells_extract])

        output_workbook.save(output_file)
        
    # Handle exceptions
    except FileNotFoundError as fnf_error:
        print(f"Error: {fnf_error}. Please enter a valid username")

    except Exception as e:
        print(f"An error occured: {e}")


if __name__ == "__main__":
    main()
