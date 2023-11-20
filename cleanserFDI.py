import pandas as pd

import_file_name = "Original FDI Data by NAICS.xlsx"
export_file_name = "Cleaned FDI Data by NAICS.xlsx"

df = pd.read_excel(import_file_name, skiprows=10, skipfooter=22)
df = df.drop(index=0)
df = df.reset_index()
df.rename(columns={'North American Industry Classification System (NAICS) 2 3 4' : 'NAICS'}, inplace=True)
naics = 'nan'

# Add NAICS sector names for foreign investment rows
for index, row in df.iterrows():
    naics = row['NAICS'] if not index % 2 else naics
    if index % 2:
        df.at[index, 'NAICS'] = naics

df = df.filter(items=[i for i in range(len(df)) if i % 2], axis=0) # Filter out rows of canadian investment abroad
df = df.reset_index()

df = df.drop(['level_0', 'index', 'Canadian and foreign direct investment'], axis=1) # Drop unecessary indices

# Clean data cells
for row_i, row in df.iterrows(): 
    for col_i, col in enumerate(row):
        if not col_i and ',' in col:
            for i, c in enumerate(col):
                if c == ',' and col[i+1] != ' ':
                    df.iat[row_i, col_i] = col[:i+1] + " " + col[i+1:]

        if col == ".." or col == "x":
            df.iat[row_i, col_i] = None
        if col_i and type(col) == type("") and 't' in col:
            cell = col.replace('t', '')
            cell = cell.replace(',', '')
            df.iat[row_i, col_i] = int(cell)

df.to_excel(export_file_name)