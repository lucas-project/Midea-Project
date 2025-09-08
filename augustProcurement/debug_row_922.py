import pandas as pd

# Read the Excel file
df = pd.read_excel('八月采购.xlsx')
df.columns = ['A', 'B', 'C']

print("=== DEBUGGING ROW 922 ===")

# Check row 922 specifically
row_922_b = str(df.iloc[922]['B']) if pd.notna(df.iloc[922]['B']) else 'NaN'
row_922_c = str(df.iloc[922]['C']) if pd.notna(df.iloc[922]['C']) else 'NaN'

print(f"Row 922: B='{row_922_b}' | C='{row_922_c}'")

# Extract product name from the Total line
if 'Total:' in row_922_b:
    product_name = row_922_b.split('Total:')[0].strip()
    print(f"Extracted product name: '{product_name}'")
    print(f"Product name length: {len(product_name)}")
    print(f"Product name is empty: {product_name == ''}")
    print(f"Product name stripped: '{product_name.strip()}'")
    print(f"Product name stripped is empty: {product_name.strip() == ''}")

# Check the context around row 922
print("\n=== CONTEXT AROUND ROW 922 ===")
for i in range(915, 925):
    if i < len(df):
        b_val = str(df.iloc[i]['B']) if pd.notna(df.iloc[i]['B']) else 'NaN'
        c_val = str(df.iloc[i]['C']) if pd.notna(df.iloc[i]['C']) else 'NaN'
        marker = " <-- HERE" if i == 922 else ""
        print(f"  Row {i}: B='{b_val}' | C='{c_val}'{marker}")



