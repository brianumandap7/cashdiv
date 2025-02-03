import pandas as pd

# Load the Excel file
file_path = 'num.xlsx'  # Replace with your actual file name
df = pd.read_excel(file_path, header=None, dtype=str)

# Remove the decimal point
df[0] = df[0].str.replace('.', '', regex=False)

# Save the modified data to a new Excel file without the apostrophe
output_file = 'cleanB.xlsx'
df.to_excel(output_file, index=False, header=False)

print(f"Modified file saved as {output_file}")
