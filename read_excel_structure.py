import pandas as pd

# Read the Excel file
excel_file = 'sample/Report.xlsx'
df = pd.read_excel(excel_file)

# Display column information
print("\nColumns in the Excel file:")
for idx, col in enumerate(df.columns):
    print(f"{idx}: {col} (Type: {df[col].dtype})")

print("\nFirst 10 rows with all columns:")
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
print(df.head(10))

print("\nChecking 'To' column contents:")
# Get the index of the 'To' column
to_column = None
for idx, col in enumerate(df.columns):
    if str(col).strip() == 'To':
        to_column = idx
        break

if to_column is not None:
    print("\nFirst 10 'To' column entries:")
    for idx, value in enumerate(df.iloc[:10, to_column]):
        print(f"{idx}: {value}")
else:
    print("No 'To' column found in the Excel file")
