
import xlwings as xl
import pandas as pd

xlfile = r"C:\Users\john\OneDrive\Documents\NewArtistReceiptsPrototype.xlsm"
wb = xl.Book(xlfile)
sheet = wb.sheets('ReceiptSheet')
table = sheet.tables('ReceiptsTable')
rng = table.range
df = pd.DataFrame(rng.value)

# Make the dataframe recognise the tow row as column headers
df.columns = df.iloc[0]
df = df[1:]
df = df.reset_index(drop=True)

# Group the DataFrame by 'Artist' and sum the 'Total' values n
artist_totals = df.groupby('Artist')['Total'].sum().reset_index()

# Sort the DataFrame by 'Artist' in ascending order
artist_totals_sorted = artist_totals.sort_values(by='Total', ascending=False)

# Display the DataFrame with total sums alongside each artist
print(artist_totals_sorted)

