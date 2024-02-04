import pandas as zx
import glob

filepath = glob.glob("invoices/*.xlsx")

for i in filepath:
    data = zx.read_excel(i, sheet_name="Sheet 1")
    print(data)
