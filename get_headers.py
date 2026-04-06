import pandas as pd
df = pd.read_excel('Master_Orders.xlsx')
with open('cols.txt', 'w') as f:
    f.write(str(df.columns.tolist()))
