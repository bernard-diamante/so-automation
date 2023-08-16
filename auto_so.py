import tkinter as tk
from tkinter import filedialog
from ctypes import windll
from openpyxl import Workbook, load_workbook

from excel_manip import create_directory, convert_xls_to_xlsx, auto_size_columns, duplicate_excel_file, set_list_of_pivot_tables_refresh_on_load
from pop_raw import populate_raw_data_sheet


def select_file():
    global output_file_path
    output_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if output_file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, output_file_path)

def submit():
    if "output_file_path" in globals() and output_file_path != "":
        app.destroy()  # Close the GUI window and continue with the program
    else:
        pass

def create_gui(gui_title):
    global file_entry
    global app  # Declare app as a global variable
    app = tk.Tk()
    app.geometry("700x350")
    app.title(gui_title)
    
    # Create GUI components here
    file_label = tk.Label(app, text="Select a file:")
    file_label.grid(row=0, column=0, columnspan=2, pady=10)

    file_entry = tk.Entry(app, width=50)
    file_entry.grid(row=1, column=0, columnspan=2, pady=5)

    browse_button = tk.Button(app, text="Browse", command=select_file)
    browse_button.grid(row=2, column=0, padx=5, pady=10, sticky="e")  # Align to the east (right)

    submit_button = tk.Button(app, text="Submit", command=submit)
    submit_button.grid(row=2, column=1, padx=5, pady=10, sticky="w")  # Align to the west (left)

    app.mainloop()

def main():
    try:
        # windll.shcore.SetProcessDpiAwareness(1)

        # Initialize the GUI
        # create_gui("Create Service Overview")
        # app.mainloop()
        
        xls_dir = "xls"
        xlsx_dir = "xlsx"
        template_file_path = "SO_Template.xlsx"
        output_file_name = "Service Overview.xlsx"
        output_file_path = output_file_name

        # Convert downloaded .xls to .xlsx Service files
        create_directory("xlsx")
        service_files = convert_xls_to_xlsx(xls_dir, xlsx_dir)
        
        # Duplicate workbook as output file
        duplicate_excel_file(template_file_path, output_file_path)
        
        # Populate "raw" sheet
        output_workbook = populate_raw_data_sheet(output_file_path, service_files, xlsx_dir)
        raw_sheet = output_workbook["raw"]
        auto_size_columns(raw_sheet)

        # Set the refreshOnLoad attribute = True for all pivot tables in the workbook
        set_list_of_pivot_tables_refresh_on_load(output_file_path)

        
    # Handle exceptions
    except FileNotFoundError as fnf_error:
        print(f"Error: {fnf_error}. Please enter a valid username")

    except Exception as e:
        print(f"An error occured: {e}")

if __name__ == "__main__":
    main()
