import pandas as pd

# File paths
file_path = "demo.xlsx"  # Replace with the file path
output_file = "updated_employee_sheet.xlsx"  # Output file path

# Load both sheets
sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')  # First sheet with employee names
sheet2 = pd.read_excel(file_path, sheet_name='Sheet1 (2)')  # Second sheet with employee names and departments

# Merge the data based on the employee name
merged_df = sheet1.merge(sheet2, on='Name', how='left')  # Adjust column name if needed

# Save the updated sheet
merged_df.to_excel(output_file, index=False)
print(f"Updated file saved as {output_file}")

