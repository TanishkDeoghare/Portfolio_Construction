import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Load the data
expected_stress_loss_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Expected_Stress_loss.xlsx"
correlation_matrix_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Correlation_Matrix_Filtered.xlsx"
expected_returns_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Expected_Returns.xlsx"

# Read the expected stress loss data and drop rows with any null values
expected_stress_loss_df = pd.read_excel(expected_stress_loss_path, index_col=0, header=1)

# Read the correlation matrix data
correlation_matrix_df = pd.read_excel(correlation_matrix_path, index_col=0)

# Read the expected returns data
expected_returns_df = pd.read_excel(expected_returns_path, index_col=0, header=1)

# Drop the first three rows (Instrument, Sector, Rating) from expected_stress_loss_df
expected_stress_loss_df = expected_stress_loss_df.iloc[2:, :]

# Ensure only numeric columns are used
expected_stress_loss_df = expected_stress_loss_df.apply(pd.to_numeric, errors='coerce').dropna(axis=1)

# Print shape after processing
print("Expected Stress Loss DF shape:", expected_stress_loss_df.shape)

# Ensure the correlation matrix only includes the instruments present in expected_stress_loss_df
filtered_correlation_matrix_df = correlation_matrix_df.loc[expected_stress_loss_df.columns, expected_stress_loss_df.columns]

# Print shape after filtering correlation matrix
print("Filtered Correlation Matrix DF shape:", filtered_correlation_matrix_df.shape)

# Convert dataframes to numpy arrays for matrix operations
expected_stress_loss = expected_stress_loss_df.to_numpy()
correlation_matrix = filtered_correlation_matrix_df.to_numpy()

# Perform the matrix multiplication for each column
results = []
for column in expected_stress_loss:
    column_vector = column.reshape(-1, 1)
    transposed_vector = column_vector.T
    intermediate_result = np.matmul(correlation_matrix, column_vector)
    final_result = np.sqrt(np.matmul(transposed_vector, intermediate_result))
    results.append(final_result[0][0])

# Calculate the sum of each row in the expected returns, excluding the first three rows
expected_returns_sum = expected_returns_df.iloc[2:, :].sum(axis=1)

# Print lengths to debug
print("Length of Expected Returns Sum:", len(expected_returns_sum))
print("Length of Stress Loss Results:", len(results))

# Ensure the lengths are the same
assert len(expected_returns_sum) == len(results), "Length mismatch between expected returns and stress loss results."

# Convert results to a DataFrame
results_df = pd.DataFrame({
    'Expected_Returns': expected_returns_sum.values,
    'Stress_Loss_Result': results
})

# Save the results to Excel
output_path = r"C:/Users/tanishk.deoghare/OneDrive - Angel Oak Capital Advisors/Desktop/Portfolio Construction/Dry Run 1/Results.xlsx"
results_df.to_excel(output_path, index=False)

# Plot the graph of stress loss vs expected returns
plt.figure(figsize=(10, 6))
plt.scatter(results_df['Stress_Loss_Result'], results_df['Expected_Returns'], color='blue', alpha=0.6, label='Portfolios')

# Plot the efficient frontier for discrete values of stress loss
discrete_stress_losses = np.linspace(min(results_df['Stress_Loss_Result']), max(results_df['Stress_Loss_Result']), 10)
efficient_frontier = []

for stress_loss in discrete_stress_losses:
    max_return = results_df[results_df['Stress_Loss_Result'] <= stress_loss]['Expected_Returns'].max()
    efficient_frontier.append((stress_loss, max_return))

efficient_frontier = np.array(efficient_frontier)
plt.plot(efficient_frontier[:, 0], efficient_frontier[:, 1], color='red', marker='o', linestyle='-', label='Efficient Frontier')

plt.title('Stress Loss vs Expected Returns')
plt.xlabel('Stress Loss')
plt.ylabel('Expected Returns')
plt.grid(True)
plt.legend()
plt.show()

print("Matrix multiplication and calculations are complete. Results saved to:", output_path)
