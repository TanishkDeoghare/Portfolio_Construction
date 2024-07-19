import pandas as pd
import os

# Constant file name and sheet name
file_name = "Non-Agency RMBS - Weekly Spreads Report 2024.07.12.xlsx"
sheet_name = "Non-Agency RMBS Spreads"

# Input folder path
folder_path = "C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1"

# Combine folder path and file name to get the full file path
file_path = os.path.join(folder_path, file_name)

# Define the columns to read by their indices (DA to DV are 104 to 126, FN to FT are 169 to 175, the second column is 1)
columns_to_read = [1] + list(range(105, 126)) + list(range(169, 176))  # 0-indexed for pandas

# Read the first four rows to concatenate their text
first_four_rows = pd.read_excel(file_path, sheet_name=sheet_name, usecols=columns_to_read, skiprows=2, nrows=3, header=None)

# first_five_rows = pd.read_excel(file_path, sheet_name=sheet_name, usecols=columns_to_read, nrows=4, header=None)

# Concatenate the text from the first four rows
concatenated_row = [first_four_rows.iloc[0, 0]] + [' '.join(first_four_rows[col].dropna().astype(str)) for col in first_four_rows.columns[1:]]

# Read the remaining data, skipping the first 9 rows
df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=columns_to_read, skiprows=1, header=None)

# Create a new DataFrame with the concatenated row as the first row
new_df = pd.DataFrame([concatenated_row], columns=df.columns)
new_df = pd.concat([new_df, df], ignore_index=True)
new_df = new_df.drop([1,2,3,4])
new_df.columns = new_df.iloc[0]

new_columns = ['Date'] + [
    f"Prime 2.0 {col}" if col in new_df.columns[1:len(range(105, 126))+1] else f"Non-QM {col}" if col in new_df.columns[len(range(105, 126))+1:] else col
    for col in new_df.columns[1:]
]
new_df.columns = new_columns

new_df = new_df.drop([0])
new_df = new_df.reset_index(drop=True)

columns_to_keep = ['Date'] + [col for col in new_df.columns if 'Spread' in col or 'Non-QM' in col]
filtered_df = new_df[columns_to_keep]

# filtered_df.columns = ["Wells" if i != 0 else filtered_df.columns[i] for i in range(len(filtered_df.columns))]

# Output file path
output_file_path = os.path.join(folder_path, "Filtered_WF_Spreads_Report.xlsx")

# Save the new DataFrame to an Excel file
# new_df.to_excel(output_file_path, index=False)
filtered_df.to_excel(output_file_path, index=False)

# Display the output file path
print(f"The cleaned data has been saved to: {output_file_path}")
