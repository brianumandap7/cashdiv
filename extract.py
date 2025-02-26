import pandas as pd

def process_excel(input_file, output_file):
    # Read the Excel file without headers
    df = pd.read_excel(input_file, header=None, dtype=str)

    # Assign column names assuming column B contains the numbers
    df.columns = ["A", "B"]  

    # Extract last 8 digits into column D
    df["D"] = df["B"].str[-8:]

    # Extract the remaining part into column C
    df["C"] = df["B"].str[:-8]

    # Format column C as a number with 2 decimal places, remove the decimal point, and keep leading zeros
    df["E"] = df["C"].apply(lambda x: f"{int(float(x)*100):016d}")

    # Combine columns E and D into column F
    df["F"] = df["E"] + df["D"]

    # Create column G by removing exactly one leading zero from column F
    df["G"] = df["F"].apply(lambda x: x[1:] if x.startswith("0") else x)

    # Save to new Excel file
    df.to_excel(output_file, index=False)
    print(f"Processed file saved as {output_file}")

# Example usage
input_file = "testfeb20.xlsx"  # Update with actual file name
output_file = "extracted.xlsx"  # Desired output file name
process_excel(input_file, output_file)
