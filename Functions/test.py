import os
import path as t


import pandas as pd

data = pd.read_excel('../Data/html_data_Sales_and_Stock.xlsx', index=False)
print(data)

if data.empty:
    print('No Records')
else:
    print(data)