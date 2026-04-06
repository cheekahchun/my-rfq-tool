import pandas as pd
import os

MASTER_FILE = 'Master_Orders.xlsx'
if os.path.exists(MASTER_FILE):
    df = pd.read_excel(MASTER_FILE)
    print("Columns:", df.columns.tolist())
    print("-" * 30)
    # Check for UID 1380
    record = df[df['UID'].astype(str) == '1380']
    if not record.empty:
        print("Record found for UID 1380:")
        print(record[['UID', 'Customer Name', 'Item Name']].to_dict('records'))
    else:
        print("UID 1380 NOT found in Master_Orders.xlsx")
        print("Last 3 rows:")
        print(df.tail(3)[['UID', 'Customer Name', 'Item Name']].to_dict('records'))
else:
    print("Master_Orders.xlsx not found")
