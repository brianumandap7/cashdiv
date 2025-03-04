import pandas as pd
import re

def process_text_file(input_file, column_a_name, column_b_name):
    data = []
    
    with open(input_file, 'r', encoding='utf-8', errors='ignore') as file:
        for line in file:
            # Extract first 10 characters as ID, then extract name until digits start
            match = re.match(r'(\d{10}[A-Z][^\d]+)(\d+)', line)
            if match:
                column_a = match.group(1).strip()
                column_b = match.group(2).strip()
                data.append([column_a, column_b])
    
    return pd.DataFrame(data, columns=[column_a_name, column_b_name])

# Process both files
df1 = process_text_file('cashdiv.txt', 'Column A', 'Column B')
df2 = process_text_file('issd.txt', 'Column C', 'Column D')

# Combine both dataframes
final_df = pd.concat([df1, df2], axis=1)

# Writing to Excel
final_df.to_excel('output.xlsx', index=False)
