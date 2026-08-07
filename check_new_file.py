import pandas as pd

# Read the new Excel file
df = pd.read_excel(r'c:\Users\Lenovo\Downloads\Telegram Desktop\orders-2025-10-14-17-12-00.xlsx')

print(f'Total records: {len(df)}')
print(f'Roll number range: {df["Exam Roll"].min()} to {df["Exam Roll"].max()}')
print(f'\nColumns available:')
for col in df.columns:
    print(f'  {{{{{col}}}}}')

print(f'\nSample data (first 3 rows):')
print(df.head(3).to_string())

print(f'\nLast 3 rows:')
print(df.tail(3).to_string())
