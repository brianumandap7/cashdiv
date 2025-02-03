import pandas as pd

# Load the Excel file
file_path = 'cleanA.xlsx'  # Replace with your actual file name
df = pd.read_excel(file_path, header=None)

# Combine the values of the first and second columns into the first column
df[0] = df[0].astype(str) + ' ' + df[1]

# Save the modified data to a new Excel file
output_file = 'pay_out.xlsx'
df.to_excel(output_file, index=False, header=False)

print(f"Modified file saved as {output_file}")
