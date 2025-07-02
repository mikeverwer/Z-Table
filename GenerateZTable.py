import pandas as pd
from scipy.stats import norm

# Generate Z-values from -3.99 to 3.99 in 0.01 increments
z_values = [round(z * 0.01, 2) for z in range(-399, 400)]
probabilities = [norm.cdf(z) for z in z_values]

# Create DataFrame and save to CSV
df = pd.DataFrame({"z_value": z_values, "probability": probabilities})
df.to_csv("full_z_table.csv", index=False, float_format="%.6f")