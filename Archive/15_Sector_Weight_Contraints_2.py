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

# Function to format values as percentages
def format_as_percentage(df):
    return df.applymap(lambda x: f"{x:.2%}")

# Process each ALLOC file
for alloc_file in alloc_files:
    file_path = os.path.join(base_path, alloc_file)
    
    # Load the Excel file
    df = pd.read_excel(file_path)
    
    # Identify random number columns
    random_columns = [col for col in df.columns if 'Random_Number' in col]
    
    # Initialize a dictionary to store sector weights for all scenarios
    sector_weights_all_scenarios = {}
    
    # Calculate sector weights for each random number column
    for random_col in random_columns:
        sector_sum = df.groupby('Sector')[random_col].sum()
        total_sum = sector_sum.sum()
        sector_weights = sector_sum / total_sum
        sector_weights_all_scenarios[random_col] = sector_weights
    
    # Convert the dictionary to a DataFrame
    sector_weights_df = pd.DataFrame(sector_weights_all_scenarios)
    
    # Calculate the minimum and maximum allocations for each sector across the sector_weights_df
    min_allocations = sector_weights_df.min(axis=1)
    max_allocations = sector_weights_df.max(axis=1)
    
    min_max_df = pd.DataFrame({
        'Min': min_allocations,
        'Max': max_allocations
    })
    
    # Format the DataFrames as percentages
    # sector_weights_df = format_as_percentage(sector_weights_df)
    # min_max_df = format_as_percentage(min_max_df)
    
    # Generate the output file name
    output_file_name = f'Sector_Weights_All_Scenarios_{alloc_file.split("_")[-1].split(".")[0]}.xlsx'
    output_file_path = os.path.join(output_base_path, output_file_name)
    
    # Save the sector weights and min/max allocations to separate sheets in the same Excel file
    with pd.ExcelWriter(output_file_path) as writer:
        sector_weights_df.to_excel(writer, sheet_name='Sector Weights')
        min_max_df.to_excel(writer, sheet_name='Min_Max_Allocations')
    
    print(f"Sector weights and min/max allocations for {alloc_file} saved to {output_file_path}")
    
    # Identify columns where sector allocation values fall beyond the limits
    out_of_bounds_columns = []
    
    for random_col in random_columns:
        sector_allocation = sector_weights_df[random_col].astype(float)
        
        for sector in sector_allocation.index:
            alloc = alloc_file.split('_')[-1].split('.')[0]
            min_limit = df_controls[(df_controls['ALLOC'] == alloc) & (df_controls['Sector'] == sector)]['Min'].values[0]
            max_limit = df_controls[(df_controls['ALLOC'] == alloc) & (df_controls['Sector'] == sector)]['Max'].values[0]
            
            if sector_allocation[sector] < min_limit or sector_allocation[sector] > max_limit:
                out_of_bounds_columns.append(random_col)
                break
    
    if out_of_bounds_columns:
        # Save the out-of-bounds columns to a new Excel file
        out_of_bounds_df = sector_weights_df[out_of_bounds_columns]
        out_of_bounds_output_file_name = f'Out_of_Bounds_Sector_Weights_{alloc}.xlsx'
        out_of_bounds_output_file_path = os.path.join(output_base_path, out_of_bounds_output_file_name)
        
        with pd.ExcelWriter(out_of_bounds_output_file_path) as writer:
            out_of_bounds_df.to_excel(writer, sheet_name='Out_of_Bounds_Weights')
        
        print(f"Out-of-bounds sector weights for {alloc_file} saved to {out_of_bounds_output_file_path}")
    else:
        print(f"No out-of-bounds sector weights for {alloc_file}.")
