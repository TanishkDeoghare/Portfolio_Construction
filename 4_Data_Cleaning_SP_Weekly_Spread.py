import pandas as pd
import os

# Constant file name and sheet name
file_name = "SP spread time series 7.12.24.xlsx"
sheet_name = "Master Table"

# Input folder path
folder_path = "C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1"

# Combine folder path and file name to get the full file path
file_path = os.path.join(folder_path, file_name)

# Read the "Master Table" sheet from the Excel file
df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=9)

# Output file path
output_file_path = os.path.join(folder_path, "Cleaned_BoFA_Master_Table.xlsx")

# Save the DataFrame to an Excel file
df.to_excel(output_file_path, index=False)

# Display the output file path
print(f"The cleaned data has been saved to: {output_file_path}")
