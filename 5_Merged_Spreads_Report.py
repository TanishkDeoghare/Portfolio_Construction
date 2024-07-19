import pandas as pd
import os

# File paths
file_path_bofa = "C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Cleaned_BoFA_Master_Table.xlsx"
file_path_wf = "C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Filtered_WF_Spreads_Report.xlsx"

# Load the Excel files into DataFrames
df_bofa = pd.read_excel(file_path_bofa)
df_wf = pd.read_excel(file_path_wf)

# Merge the DataFrames on the "Date" column
merged_df = pd.merge(df_bofa, df_wf, on="Date", how="inner")

# Ensure the "Date" column is in datetime format
merged_df['Date'] = pd.to_datetime(merged_df['Date'])

# Filter the DataFrame to keep only the rows with dates up to 2018
filtered_merged_df = merged_df[merged_df['Date'].dt.year >= 2018]

# Output file path for the merged DataFrame
output_file_path = "C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Merged_Spreads_Report.xlsx"

# Save the merged DataFrame to an Excel file
# merged_df.to_excel(output_file_path, index=False)

# Display the output file path
# print(f"The merged data has been saved to: {output_file_path}")

# Save the filtered and merged DataFrame to an Excel file
filtered_merged_df.to_excel(output_file_path, index=False)

# Display the output file path
print(f"The filtered and merged data has been saved to: {output_file_path}")