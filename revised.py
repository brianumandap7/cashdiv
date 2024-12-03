from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os

def select_excel_file():
    """Prompts the user to select an Excel file and returns the file path."""
    Tk().withdraw()  # Hide the root tkinter window
    print("Opening file dialog to select an Excel file...")
    file_path = askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if not file_path:
        print("No file selected. Exiting.")
        exit()  # Exit if no file is selected
    print(f"Selected file: {file_path}")
    return file_path

def format_row(row, total_width_before_last_column, blank_spaces):
    """Formats a single row based on specifications."""
    col1 = str(row[0]) if row[0] is not None else ""
    col2 = str(row[1]) if row[1] is not None else ""
    combined_columns = (col1 + " " + col2).ljust(total_width_before_last_column)[:total_width_before_last_column]
    other_columns = " ".join([str(cell).strip() if cell is not None else "" for cell in row[2:-1]]) if len(row) > 2 else ""
    last_column = str(row[-1]).strip() if len(row) > 1 else ""
    formatted_row = combined_columns + " " + other_columns + " " + last_column
    return formatted_row

def process_excel_and_save():
    excel_file_path = select_excel_file()
    wb = load_workbook(excel_file_path)
    sheet = wb.active

    replacement_string = '9999999999DEPARTMENT OF TRANSPORTAT000001079200266000010633854605905300000 00000'
    blank_spaces = "       "  # 7 blank spaces
    total_width_before_last_column = 48

    all_data = []
    rows = list(sheet.iter_rows(values_only=True))
    for row in rows:
        formatted_row = format_row(row, total_width_before_last_column, blank_spaces)
        all_data.append(formatted_row)

    # Append the replacement string at the very bottom
    all_data.append(replacement_string)

    # Save the output
    txt_file_path = os.path.splitext(excel_file_path)[0] + "_output.txt"
    with open(txt_file_path, 'w') as file:
        file.write("\n".join(all_data))

    print(f"Processed data saved to {txt_file_path}")

if __name__ == "__main__":
    process_excel_and_save()
