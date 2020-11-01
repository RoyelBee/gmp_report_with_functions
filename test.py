import pandas as pd

df = pd.read_excel('./Data/gpm_data.xlsx')
data = df[['BRAND', 'MTD Sales Target', 'Actual Sales MTD', 'MTD Sales Acv']]
print(data)

# Bar chart : Brand Name in x-axis and 'Actual Sales MTD' in y-axis
# Line Chart: Brand Name in x-axis and 'MTD Sales Target' in y-axis
# Scatter   : Brand Name in x-axis and 'MTD Sales Acv' in y-axis