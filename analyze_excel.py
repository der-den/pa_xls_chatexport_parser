import pandas as pd

def analyze_excel(excel_file):
    # Read the Excel file
    df = pd.read_excel(excel_file, na_values=[''], keep_default_na=False)
    
    # Print column names and indices
    print("\nColumn names and indices:")
    for idx, col in enumerate(df.columns):
        print(f"Column {idx}: {col}")
    
    # Try to find specific columns we're interested in
    print("\nSearching for specific columns:")
    columns_to_find = ['Body', 'Attachment #1', 'Status', 'sender_name']
    for col in columns_to_find:
        if col in df.columns:
            idx = df.columns.get_loc(col)
            print(f"Found '{col}' at index {idx}")
        else:
            print(f"Column '{col}' not found!")
    
    # Print first row of data for verification
    print("\nFirst row data:")
    first_row = df.iloc[0]
    for col in df.columns:
        print(f"{col}: {first_row[col]}")

if __name__ == "__main__":
    excel_file = 'sample/Report.xlsx'
    analyze_excel(excel_file)
