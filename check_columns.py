import pandas as pd

# Read the Excel file with correct header row
df = pd.read_excel(r'c:\Users\Lenovo\Downloads\Telegram Desktop\ASP Online Reg. Boy.xlsx', header=1)

print('Available columns for placeholders in your Word template:\n')
for col in df.columns:
    print(f'  {{{{{col}}}}}')

print('\n' + '='*60)
print('Sample data from first record:')
print('='*60)
first_row = df.iloc[0]
for col in df.columns:
    print(f'{col}: {first_row[col]}')
