import pandas as pd
import json

df = pd.read_excel('RENDICION Nº 39 -CAJA CHICA- LITORAL.xlsx', sheet_name=0, header=None)

# Fill na with empty string
df = df.fillna('')
records = []
for i in range(25):
    records.append(df.iloc[i].tolist())

with open('excel_dump.json', 'w', encoding='utf-8') as f:
    json.dump(records, f, ensure_ascii=False, indent=2)
