import pandas as pd
import os

# Define the base path for your files
base_path = r'C:\Users\tanishk.deoghare\OneDrive - Angel Oak Capital Advisors\Desktop\Portfolio Construction\Dry Run 1'
output_base_path = r'C:\Users\tanishk.deoghare\OneDrive - Angel Oak Capital Advisors\Desktop\Portfolio Construction\Dry Run 1'

# List of ALLOC files
alloc_files = [
    'scaled_random_numbers_sector_allocations_ASCIX.xlsx',
    'scaled_random_numbers_sector_allocations_MBS.xlsx',  # Replace with actual file names
    'scaled_random_numbers_sector_allocations_AOUIX.xlsx',  # Replace with actual file names
    'scaled_random_numbers_sector_allocations_ANGLX.xlsx'   # Replace with actual file names
]

# Load the input controls to get min and max limits
input_controls_path = r'C:\Users\tanishk.deoghare\OneDrive - Angel Oak Capital Advisors\Desktop\Portfolio Construction\Master code\Input_controls.xlsx'
input_sheet_name = 'Fund_Sector'
df_controls = pd.read_excel(input_controls_path, sheet_name=input_sheet_name)

# Function to check for violations
def check_violations(sector_weights_df, df_controls, alloc):
    violations = {}
    for sector in sector_weights_df.index:
        min_limit = df_controls[(df_controls['ALLOC'] == alloc) & (df_controls['Sector'] == sector)]['Min'].values[0]
        max_limit = df_controls[(df_controls['ALLOC'] == alloc) & (df_controls['Sector'] == sector)]['Max'].values[0]
        for col in sector_weights_df.columns:
            weight = sector_weights_df.loc[sector, col]
            if weight < min_limit or weight > max_limit:
                if col not in violations:
                    violations[col] = []
                violations[col].append((sector, weight, min_limit, max_limit))
    return violations

# Process each ALLOC file
for alloc_file in alloc_files:
    file_path = os.path.join(base_path, alloc_file)
    
    # Load the "Sector Weights" sheet from the specified Excel file
    sector_weights_df = pd.read_excel(file_path, sheet_name='Sector Weights', index_col=0)
    
    # Extract the ALLOC value from the file name
    alloc = alloc_file.split('_')[-1].split('.')[0]
    
    # Check for violations
    violations = check_violations(sector_weights_df, df_controls, alloc)
    
    # Print the violations and save the results
    if violations:
        print(f"Violations found in {file_path}:")
        for col, sectors in violations.items():
            print(f"Column: {col}")
            for sector, weight, min_limit, max_limit in sectors:
                print(f"  Sector: {sector}, Weight: {weight:.4f}, Min: {min_limit:.4f}, Max: {max_limit:.4f}")
        
        # Collect the columns with violations
        violation_columns = list(violations.keys())
        
        # Extract the columns with violations
        violations_df = sector_weights_df[violation_columns]
        
        # Save the violations to a new Excel file
        output_file_name = f'Violations_Sector_Weights_{alloc}.xlsx'
        output_file_path = os.path.join(output_base_path, output_file_name)
        with pd.ExcelWriter(output_file_path) as writer:
            violations_df.to_excel(writer, sheet_name='Violations')
            for col, sectors in violations.items():
                detail_df = pd.DataFrame(sectors, columns=['Sector', 'Weight', 'Min', 'Max'])
                detail_df.to_excel(writer, sheet_name=col)
        
        print(f"Violations saved to {output_file_path}")
    else:
        print(f"No violations found in {file_path}.")
