import pandas as pd
import os

# Define the paths for the input files
sector_weights_file_path = r'C:\Users\tanishk.deoghare\OneDrive - Angel Oak Capital Advisors\Desktop\Portfolio Construction\Dry Run 1\Sector_Weights_All_Scenarios_MBS.xlsx'
input_controls_path = r'C:\Users\tanishk.deoghare\OneDrive - Angel Oak Capital Advisors\Desktop\Portfolio Construction\Master code\Input_controls.xlsx'
input_sheet_name = 'Fund_Sector'

# Load the input controls to get min and max limits
df_controls = pd.read_excel(input_controls_path, sheet_name=input_sheet_name)

# Load the "Sector Weights" sheet from the specified Excel file
sector_weights_df = pd.read_excel(sector_weights_file_path, sheet_name='Sector Weights', index_col=0)

# Function to check for violations
def check_violations(sector_weights_df, df_controls, alloc):
    violations = []
    for sector in sector_weights_df.index:
        min_limit = df_controls[(df_controls['ALLOC'] == alloc) & (df_controls['Sector'] == sector)]['Min'].values[0]
        max_limit = df_controls[(df_controls['ALLOC'] == alloc) & (df_controls['Sector'] == sector)]['Max'].values[0]
        for col in sector_weights_df.columns:
            weight = sector_weights_df.loc[sector, col]
            if weight < min_limit or weight > max_limit:
                violations.append((sector, col, weight, min_limit, max_limit))
    return violations

# Extract the ALLOC value from the file name
alloc = "MBS"

# Check for violations
violations = check_violations(sector_weights_df, df_controls, alloc)

# Print the violations
if violations:
    print(f"Violations found in {sector_weights_file_path}:")
    for sector, col, weight, min_limit, max_limit in violations:
        print(f"Sector: {sector}, Column: {col}, Weight: {weight:.4f}, Min: {min_limit:.4f}, Max: {max_limit:.4f}")
    
    # Collect the columns with violations
    violation_columns = list(set([col for _, col, _, _, _ in violations]))
    
    # Extract the columns with violations
    violations_df = sector_weights_df[violation_columns]
    
    # Save the violations to a new Excel file
    output_file_path = r'C:\Users\tanishk.deoghare\OneDrive - Angel Oak Capital Advisors\Desktop\Portfolio Construction\Dry Run 1\Violations_Sector_Weights_MBS.xlsx'
    violations_df.to_excel(output_file_path, sheet_name='Violations')
    
    print(f"Violations saved to {output_file_path}")
else:
    print(f"No violations found in {sector_weights_file_path}.")
