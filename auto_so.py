import os
from openpyxl import Workbook, load_workbook
import pandas as pd


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
        workbook = load_workbook(file_path)
        sheet = workbook.active
        return sheet[cell].value
    except Exception as e:
        print(f"Error: {e}")

def main():
    try:
        username = str(input("Please input your username: "))
        desktop_path = f"C:\\Users\\{username}\\Desktop\\"
        input_folder = f"{desktop_path}ExcelFiles"
        output_folder = f"{desktop_path}ExcelFiles_Converted"
        output_file = "Service Overview.xlsx"
        output_sheet = "raw data"

        # Field list
        cells_to_extract = {"SERVICE DESC": "D3", "FIRST VESSEL ON LIST": "C8"}
        headers_list = [cell for cell in cells_to_extract.values()]
        
        # .xlsx Service files
        service_files = convert_xls_to_xlsx(input_folder, output_folder)
        
        # Create output Service Overview file with headers
        output_workbook = Workbook()
        output_workbook.remove(output_workbook.active)
        output_sheet = output_workbook.create_sheet(output_sheet)
        output_sheet.append([key for key in cells_to_extract.keys()])

        for service_file in service_files:
            file_path = os.path.join(output_folder, service_file)
            output_sheet.append([extract_cell(file_path, cell) for cell in headers_list])

        output_workbook.save(output_file)
        
    # Handle exceptions
    except FileNotFoundError as fnf_error:
        print(f"Error: {fnf_error}. Please enter a valid username")

    except Exception as e:
        print(f"An error occured: {e}")


if __name__ == "__main__":
    main()
