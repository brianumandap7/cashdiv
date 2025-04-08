import pandas as pd

def process_excel(input_file, output_file):
    # Load the Excel file
    df = pd.read_excel(input_file, dtype={"B": str, "D": str})
    
    # Ensure column names are correctly inferred
    df.columns = ["A", "B", "C", "D", "E"]
    
    # Process only rows where column E is "Not Equal"
    def process_row(row):
        if row["E"] == "Not Equal":
            # Ensure values are strings and long enough
            b_str = str(row["B"])
            d_str = str(row["D"])
            
            if len(b_str) > 8 and len(d_str) > 8:
                # Remove last 8 digits and convert to float
                b_val = float(b_str[:-8]) / 100
                d_val = float(d_str[:-8]) / 100
                
                # Remove leading zeros
                b_val = float(str(b_val).lstrip("0"))
                d_val = float(str(d_val).lstrip("0"))
                
                # Subtract B from D
                return d_val - b_val
        return ""
    
    df["F"] = df.apply(process_row, axis=1)
    
    # Save the processed data to a new Excel file
    df.to_excel(output_file, index=False)

# File names
input_file = "comparison.xlsx"
output_file = "report_feb20.xlsx"

# Run the function
process_excel(input_file, output_file)