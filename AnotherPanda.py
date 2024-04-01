import pandas as pd
import openpyxl as op

# Load the Excel file
xlfile = r"C:\Users\john\OneDrive\Documents\NewArtistReceiptsPrototype.xlsm"
workbook = op.load_workbook(xlfile, data_only=True)

# Specify the sheet name and table name (ListObject)
sheet_name = 'ReceiptSheet'
table_name = 'ReceiptsTable'

# Access the ListObject (Table) range
#sheet = workbook[sheet_name]
#table = sheet[table_name]


#df = pd.DataFrame(table.values, columns=table.headers)

# Now 'df' contains the data from the dynamic ListObject
#print(df.head())