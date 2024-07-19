import pandas as pd
import numpy as np

# Path to the filtered mapped spreads and stress loss report
filtered_spreads_stress_loss_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Filtered_Mapped_Spreads_And_Stress_Loss_Report.xlsx"
input_controls_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Master code/Input_controls.xlsx"
input_sheet_name = 'Fund_Sector'

# Load the filtered data
df_filtered_spreads_stress_loss = pd.read_excel(filtered_spreads_stress_loss_path)
df_controls = pd.read_excel(input_controls_path, sheet_name=input_sheet_name)

# Group by 'ALLOC' and 'Sector' to find min and max limits
min_max_limits = df_controls.groupby(['ALLOC', 'Sector']).agg({'Min': 'min', 'Max': 'max'}).reset_index()

# Number of random numbers to generate per instrument
num_random_numbers = 200

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

# Generate random numbers for each instrument
np.random.seed(42)  # Set a seed for reproducibility
random_numbers = np.random.random((len(df_filtered_spreads_stress_loss), num_random_numbers))

# Create a new DataFrame with instrument names and random numbers
random_numbers_df = pd.DataFrame(random_numbers, columns=[f'Random_Number_{i+1}' for i in range(num_random_numbers)])
df_random_numbers = pd.concat([df_filtered_spreads_stress_loss[['Instrument','Sector','Ratings']], random_numbers_df], axis=1)

# Initialize a DataFrame to store scaled random numbers
scaled_random_numbers_df = df_random_numbers[['Instrument', 'Sector', 'Ratings']].copy()


# Iterate through each random allocation and scale the random numbers for each sector
for i in range(num_random_numbers):
    random_col = f'Random_Number_{i+1}'
    total_allocation_col = df_new_allocations.iloc[i::num_random_numbers]['Allocation'].values
    sector_allocations = df_random_numbers.groupby('Sector')[random_col].sum()
    
    scaled_random_col = []
    for sector in df_random_numbers['Sector'].unique():
        sector_df = df_random_numbers[df_random_numbers['Sector'] == sector]
        # scale_factor = total_allocation_col[df_new_allocations['Sector'] == sector].values[0] / sector_allocations[sector]
        scale_factor = df_new_allocations[df_new_allocations['Sector'] == sector].values[i][2]/sector_allocations[sector]
        scaled_random_col.extend(sector_df[random_col] * scale_factor)
    
    scaled_random_numbers_df[random_col] = scaled_random_col

# Print the scaled random numbers DataFrame
sector_scaled_random_numbers = scaled_random_numbers_df.iloc[:,3:]
scaled_random_numbers_df_scaled = sector_scaled_random_numbers.div(sector_scaled_random_numbers.sum(axis=0), axis=1)

concat_scaled_random_numbers_df_scaled = pd.concat([scaled_random_numbers_df[['Instrument','Sector','Ratings']], scaled_random_numbers_df_scaled], axis=1)

concat_scaled_random_numbers_df_scaled.to_excel("C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/scaled_random_numbers_sector_allocations.xlsx", index=False)