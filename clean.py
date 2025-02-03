import pandas as pd
import re

# Load the Excel file
file_path = 'acc.xlsx'  # Replace with your file name
df = pd.read_excel(file_path, header=None)

# Flatten the data to a single list
all_data = df.values.flatten()

# Filter out NaN values
all_data = [str(item) for item in all_data if pd.notnull(item)]

# Process the data
processed_data = []
for item in all_data:
    # Use regex to match the number followed by the name
    matches = re.findall(r'(\d+)([A-Z ,\.]+)', item)
    for match in matches:
        number, name = match
        name = name.strip()
        if number and name:
            processed_data.append([number, name])

# Create a new DataFrame
output_df = pd.DataFrame(processed_data, columns=['Number', 'Name'])

# Save to a new Excel file
output_file = 'cleanA.xlsx'
output_df.to_excel(output_file, index=False)

print(f'Data has been processed and saved to {output_file}')
