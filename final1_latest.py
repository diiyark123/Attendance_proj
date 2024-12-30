import pandas as pd

# File path to your input Excel file
file_path = "Daily Attendance Report (5).xls"
sheet_name = "DailyAttendance_DetailedReport"

# Read the entire sheet without setting headers initially
data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

# Initialize an empty list to store tables
tables = []

# Initialize variables for table tracking
current_date = None
current_table = []

# Iterate through rows to identify metadata and extract tables
for _, row in data.iterrows():
    # Check if "Attendance Date" metadata is in the row
    if row.dropna().str.contains("Attendance Date", case=False, na=False).any():
        # Extract the date (assuming it's the last non-empty cell in the row)
        current_date = row.dropna().to_list()[-1]

        # If a previous table exists, process it
        if current_table:
            # Create a DataFrame from the current table
            df = pd.DataFrame(current_table[1:], columns=current_table[0])

            # Delete empty rows and columns
            df.dropna(how='all', inplace=True)  # Remove rows where all values are NaN
            df.dropna(axis=1, how='all', inplace=True)  # Remove columns where all values are NaN

            # Insert the "Date" column in the second position next to "E. Code"
            if "E. Code" in df.columns:
                e_code_index = df.columns.get_loc("E. Code") + 1  # Find the position of "E. Code" and add 1
                df.insert(e_code_index, "Date", current_date)  # Insert "Date" after "E. Code"
            else:
                df["Date"] = current_date  # If "E. Code" is not found, just add "Date" at the end

            # Append the processed table
            tables.append(df)
            current_table = []  # Reset the current table
    elif row.dropna().empty:  # Skip empty rows
        continue
    else:
        # Add the current row to the current table
        current_table.append(row.to_list())

# Process the last table if it exists
if current_table:
    df = pd.DataFrame(current_table[1:], columns=current_table[0])

    # Delete empty rows and columns
    df.dropna(how='all', inplace=True)  # Remove rows where all values are NaN
    df.dropna(axis=1, how='all', inplace=True)  # Remove columns where all values are NaN

    # Insert the "Date" column in the second position next to "E. Code"
    if "E. Code" in df.columns:
        e_code_index = df.columns.get_loc("E. Code") + 1  # Find the position of "E. Code" and add 1
        df.insert(e_code_index, "Date", current_date)  # Insert "Date" after "E. Code"
    else:
        df["Date"] = current_date  # If "E. Code" is not found, just add "Date" at the end

    tables.append(df)  # Append the final table

# Save each table as a new Excel file
for i, table in enumerate(tables):
    output_path = f"Processed_Table_{i + 1}.xlsx"
    table.to_excel(output_path, index=False)
    print(f"Processed table {i + 1} has been saved to {output_path}")

