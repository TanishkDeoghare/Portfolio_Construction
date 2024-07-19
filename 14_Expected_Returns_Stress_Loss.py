import pandas as pd
import numpy as np

# Path to the filtered mapped spreads and stress loss report
filtered_spreads_stress_loss_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Filtered_Mapped_Spreads_And_Stress_Loss_Report.xlsx"

# Load the filtered data
df_filtered_spreads_stress_loss = pd.read_excel(filtered_spreads_stress_loss_path)

# Number of random numbers to generate per instrument
num_random_numbers = 2000

# Generate random numbers for each instrument
np.random.seed(42)  # Set a seed for reproducibility
random_numbers = np.random.random((len(df_filtered_spreads_stress_loss), num_random_numbers))

# Create a new DataFrame with instrument names and random numbers
random_numbers_df = pd.DataFrame(random_numbers, columns=[f'Random_Number_{i+1}' for i in range(num_random_numbers)])
df_random_numbers = pd.concat([df_filtered_spreads_stress_loss[['Instrument','Sector','Ratings']], random_numbers_df], axis=1)

# Scale the random numbers so that they sum to 1 for each instrument
random_numbers_scaled = random_numbers_df.div(random_numbers_df.sum(axis=0), axis=1)
scaled_random_numbers_df = pd.concat([df_filtered_spreads_stress_loss[['Instrument','Sector','Ratings']], random_numbers_scaled], axis=1)

# Extract total returns from the filtered data
total_returns = df_filtered_spreads_stress_loss['Total return']
stress_loss = df_filtered_spreads_stress_loss['Stress loss']


# Calculate expected returns by multiplying scaled random numbers with total returns
expected_returns = random_numbers_scaled.multiply(total_returns, axis=0)
expected_returns_df = pd.concat([df_filtered_spreads_stress_loss[['Instrument','Sector','Ratings']], expected_returns], axis=1)

expected_stress_loss = random_numbers_scaled.multiply(stress_loss, axis=0)
expected_stress_loss_df = pd.concat([df_filtered_spreads_stress_loss[['Instrument','Sector','Ratings']], expected_stress_loss], axis=1)

# Save the expected returns to an Excel file
# output_expected_returns_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Expected_Returns_verticle.xlsx"
# expected_returns_df.to_excel(output_expected_returns_path, index=False)

# output_expected_stress_loss_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Expected_Stress_loss_verticle.xlsx"
# expected_stress_loss_df.to_excel(output_expected_stress_loss_path, index=False)


# Transpose the DataFrames
expected_returns_transposed = expected_returns_df.transpose()
expected_stress_loss_transposed = expected_stress_loss_df.transpose()

# # Save the expected returns to an Excel file
output_expected_returns_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Expected_Returns.xlsx"
expected_returns_transposed.to_excel(output_expected_returns_path, index=True)

output_expected_stress_loss_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Expected_Stress_loss.xlsx"
expected_stress_loss_transposed.to_excel(output_expected_stress_loss_path, index=True)

print("Expected returns have been generated")
print(expected_returns_transposed)
