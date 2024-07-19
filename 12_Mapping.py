import pandas as pd

# Load the Excel files
spreads_and_stress_loss_path = 'C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Calculated_Spreads_And_Stress_Loss_Report.xlsx'
input_controls_path = 'C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Input_controls.xlsx'

# Load spreads and stress loss data
df_spreads_stress_loss = pd.read_excel(spreads_and_stress_loss_path)

# Load mappings data
df_mappings = pd.read_excel(input_controls_path, sheet_name='Mappings')

# Merge the data based on the 'Instrument' column
merged_df = pd.merge(df_spreads_stress_loss, df_mappings, on='Instrument', how='left')

# Save the results to an Excel file
output_path = 'Mapped_Spreads_And_Stress_Loss_Report.xlsx'
merged_df.to_excel(output_path, index=False)

print("Merged data with sector and rating:")
print(merged_df)
