import pandas as pd

# Load the Excel file
file_path = 'new.xlsx'  # Replace with your file path
df = pd.read_excel(file_path)

# Function to calculate OT Duration from Punch Records
def calculate_ot_duration(punch_records):
    if pd.isna(punch_records):
        return "00:00:00"  # Skip empty records

    # Split punch records into individual entries
    punch_entries = punch_records.split(',')
    total_duration = pd.Timedelta(0)

    # Initialize variables to track processing state
    previous_out_time = None  # To store the "out" time temporarily

    # Process punch records
    for entry in punch_entries:
        # Remove the "(TD)" suffix and strip any surrounding whitespace
        entry = entry.strip().replace("(TD)", "")

        # Split the entry into time and tag (in or out)
        if 'in' in entry:
            time = entry[:-2].strip()  # Remove the 'in' tag and extract time
            tag = 'in'
        elif 'out' in entry:
            time = entry[:-3].strip()  # Remove the 'out' tag and extract time
            tag = 'out'
        else:
            continue

        # Strip any trailing colon (:) and ensure time is properly formatted (HH:MM)
        time = time.rstrip(":")  # Remove trailing colon if present

        # Parse the time as a datetime object
        parsed_time = pd.to_datetime(time, format='%H:%M', errors='coerce')

        # If time parsing fails, skip the entry
        if pd.isna(parsed_time):
            continue

        # Process "out" and "in" pairs
        if tag == "out":  # If "out" is found
            previous_out_time = parsed_time
        elif tag == "in" and previous_out_time:  # If "in" is found after "out"
            # Calculate the time difference between 'out' and 'in'
            time_diff = parsed_time - previous_out_time
            total_duration += time_diff

            previous_out_time = None  # Reset out time after pairing

    # Convert total_duration to hours, minutes, and seconds
    total_seconds = total_duration.total_seconds()
    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    seconds = int(total_seconds % 60)

    # Return formatted duration as "HH:MM:SS"
    return f"{hours:02}:{minutes:02}:{seconds:02}"

# Calculate Work Duration as A. OutTime - A. InTime
df['Work Duration'] = df.apply(
    lambda row: str(pd.to_timedelta(row['A. OutTime']) - pd.to_timedelta(row['A. InTime']))
    .replace("0 days ", "")  # Remove "0 days" from the result
    if pd.notna(row['A. InTime']) and pd.notna(row['A. OutTime'])
    else "00:00:00",
    axis=1
)

# Calculate OT Duration
df['OT Duration'] = df['Punch Records'].apply(
    lambda x: calculate_ot_duration(x) if pd.notna(x) else "00:00:00"
)

# Function to convert HH:MM:SS strings to Timedelta
def time_to_timedelta(time_str):
    try:
        if "days" in time_str:
            return pd.Timedelta(time_str)
        hours, minutes, seconds = map(int, time_str.split(':'))
        return pd.Timedelta(hours=hours, minutes=minutes, seconds=seconds)
    except Exception:
        return pd.Timedelta(0)

# Convert 'Work Duration' and 'OT Duration' to Timedelta for calculation
df['Work Duration Timedelta'] = df['Work Duration'].apply(time_to_timedelta)
df['OT Duration Timedelta'] = df['OT Duration'].apply(time_to_timedelta)

# Calculate Total Duration as Work Duration - OT Duration
df['Total Duration Timedelta'] = df['Work Duration Timedelta'] - df['OT Duration Timedelta']

# Convert Total Duration back to HH:MM:SS format
df['Total Duration'] = df['Total Duration Timedelta'].apply(
    lambda x: str(x).split(' ')[-1] if len(str(x).split(' ')) > 1 else "00:00:00"
)

# Drop intermediate columns if not needed
drop_columns = [
    'Work Duration Timedelta', 'OT Duration Timedelta', 'Total Duration Timedelta',
    'Shift', 'E. Code', 'Late By','OT', 'Early Going By','Tot. Dur.','Work Dur.','LateBy','EarlyGoingBy', 'S. InTime', 'S. OutTime'
]

df.drop(columns=[col for col in drop_columns if col in df.columns], inplace=True)

# Reorder the columns to make 'Date' the second column after 'S.No'
columns = list(df.columns)
if 'Date' in columns:
    columns.insert(1, columns.pop(columns.index('Date')))  # Move 'Date' to the second position
    df = df[columns]

# Save the updated DataFrame back to Excel
output_file_path = 'updated_cleaned_output.xlsx'  # Replace with your desired output file path
df.to_excel(output_file_path, index=False)

print(f"Updated data has been saved to {output_file_path}")
