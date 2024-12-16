import pandas as pd

# File paths
input_file_path = r"C:\Users\AngelaTenkorang\OneDrive - AmaliTech gGmbH\Documents\progress data codecademy\Progress 13-12-2024.xlsx"
output_file_path = r"C:\Users\AngelaTenkorang\OneDrive - AmaliTech gGmbH\Documents\progress data codecademy\13_12_prog_up.xlsx"

# Sheet name
progress_sheet_name = "Progress 13-12-2024"

# Load the Excel file and sheet
try:
    data = pd.read_excel(input_file_path, sheet_name=progress_sheet_name)
    print("Columns:", data.columns)
except FileNotFoundError:
    print(f"File not found: {input_file_path}")
    exit()
except ValueError as e:
    print(f"Error loading sheet '{progress_sheet_name}': {e}")
    exit()

# Columns to remove
columns_to_remove = ["content_id", "content_slug", "content_type", "modules_completed","modules_total","created_at"]

# Check and remove columns
data_cleaned = data.drop(columns=[col for col in columns_to_remove if col in data.columns], errors='ignore')

# Split by "groups" and sort each group by "Enrollment Type"
output_sheets = {}
if "groups" in data_cleaned.columns and "enrollment_type" in data_cleaned.columns:
    groups = data_cleaned["groups"].unique()
    for group in groups:
        group_data = data_cleaned[data_cleaned["groups"] == group]
        group_data_sorted = group_data.sort_values(by="enrollment_type")
        output_sheets[str(group)] = group_data_sorted
else:
    output_sheets["Error"] = pd.DataFrame({"Message": ["Required columns missing (groups or enrollment_type)."]})

# Save to the same file with each group in a new sheet
try:
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        for sheet_name, df in output_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)  # Sheet names max 31 chars
    print(f"Updated file saved to: {output_file_path}")
except Exception as e:
    print(f"Error saving the updated file: {e}")
