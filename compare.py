import pandas as pd

def compare_columns(input_file, output_file):
    # Load the Excel file
    df = pd.read_excel(input_file, dtype=str)  # Read columns as strings to preserve leading zeros
    
    # Print column names for debugging
    print("Columns found in the file:", df.columns.tolist())
    
    # Ensure the necessary columns exist
    if df.shape[1] < 4:
        print("Error: The file does not contain enough columns.")
        return
    
    # Assuming columns B and D correspond to the second and fourth columns (index 1 and 3)
    col_b = df.columns[1]  # Second column
    col_d = df.columns[3]  # Fourth column
    
    # Compare the first 15 digits of columns B and D, starting from row 2
    def compare_values(row):
        val_b = str(row[col_b])[:15] if pd.notna(row[col_b]) else ""
        val_d = str(row[col_d])[:15] if pd.notna(row[col_d]) else ""
        return "OK" if val_b == val_d else "Not Equal"
    
    df.loc[1:, 'E'] = df.loc[1:].apply(compare_values, axis=1)
    
    # Save the output to a new Excel file
    df.to_excel(output_file, index=False)
    print(f"Comparison completed. Output saved to {output_file}")

# Run the function
compare_columns('compare.xlsx', 'comparison.xlsx')