import pandas as pd
import numpy as np

# Path to the filtered mapped spreads and stress loss report
filtered_spreads_stress_loss_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Master code/Filtered_Mapped_Spreads_And_Stress_Loss_Report.xlsx"
input_controls_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Master code/Input_controls.xlsx"
input_sheet_name = 'Fund_Sector'

# Load the filtered data
df_filtered_spreads_stress_loss = pd.read_excel(filtered_spreads_stress_loss_path)
df_controls = pd.read_excel(input_controls_path, sheet_name=input_sheet_name)

# Group by 'ALLOC' and 'Sector' to find min and max limits
min_max_limits = df_controls.groupby(['ALLOC', 'Sector']).agg({'Min': 'min', 'Max': 'max'}).reset_index()

# Number of random numbers to generate per instrument
num_random_numbers = 2000

# Generate random sector allocations
new_allocations = []

for _, row in min_max_limits.iterrows():
    alloc = row['ALLOC']
    sector = row['Sector']
    min_limit = row['Min']
    max_limit = row['Max']
    
    random_numbers = np.random.uniform(0, 1, num_random_numbers)
    sector_allocations = min_limit + (max_limit - min_limit) * random_numbers
    
    for allocation in sector_allocations:
        new_allocations.append({
            'ALLOC': alloc,
            'Sector': sector,
            'Allocation': allocation
        })

df_new_allocations = pd.DataFrame(new_allocations)

# Save the new sector allocations to an Excel file
output_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Master code/New_Sector_Allocations.xlsx"
df_new_allocations.to_excel(output_path, index=False)

print(f"New sector allocations saved to {output_path}")
