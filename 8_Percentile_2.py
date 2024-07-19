import pandas as pd
import numpy as np

# Load the Excel file
# file_path = 'Merged_Spreads_Report.xlsx'
file_path = 'Filled_Merged_Spreads_Report.xlsx'
df = pd.read_excel(file_path)

# Define the function to calculate the percentile
def calculate_percentile(series, value):
    return np.sum(series < value) / len(series) * 100

# Ensure 'Date' column is in datetime format
df['Date'] = pd.to_datetime(df['Date'])

# Extract the necessary data for each calculation
max_2020 = df[df['Date'].dt.year == 2020].max()
max_2022 = df[df['Date'].dt.year == 2022].max()
min_2021 = df[df['Date'].dt.year == 2021].min()
avg_2022_H2 = df[(df['Date'].dt.year == 2022) & (df['Date'].dt.month > 6)].mean()
avg_2021 = df[df['Date'].dt.year == 2021].mean()

# Get the latest spreads
latest_spreads = df.iloc[0]

# Calculate percentiles
percentiles = df.apply(lambda x: calculate_percentile(x, latest_spreads[x.name]) if np.issubdtype(x.dtype, np.number) else np.nan)

# Create the resulting DataFrame
result = pd.DataFrame({
    'Instrument': df.columns,  # Assuming the columns represent instruments
    'Latest Spreads': latest_spreads,
    'Covid wides': max_2020,
    '2022 wides': max_2022,
    '2021 tights': min_2021,
    'Avg 2022 H2': avg_2022_H2,
    'Avg 2021': avg_2021,
    'Percentile': percentiles
})

# Reset the index to have a column for instruments
# result.reset_index(inplace=True)
# result.rename(columns={'index': 'Instrument'}, inplace=True)

# Display the result
print(result)

# Save to Excel if needed
result.to_excel('Percentile_Spreads_Report.xlsx', index=False)
