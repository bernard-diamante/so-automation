from excel_manip import create_directory, convert_xls_to_xlsx, auto_size_columns
from pop_raw import populate_raw_data_sheet
from openpyxl import Workbook

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
        sheet = output_workbook.active
        auto_size_columns(sheet)
        
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
