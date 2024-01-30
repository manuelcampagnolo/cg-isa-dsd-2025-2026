import pandas as pd
from pathlib import Path

p=Path().absolute() # or just p=Path('.').absolute()

# file path and name
fn= p / 'ficheiros_DSD' / 'DSD_inform_202324_v3.xlsx'

# read workbook with Excel
wb = pd.ExcelFile(fn)  # funciona mesmo com ficheiro aberto em Excel
# List worksheets
print(wb.sheet_names)

# read worksheet with Pandas
sn='uc_2024-25' # sheet name
# 2 equivalent possibilities to create DataFrame:
# A: from wb
dfA=wb.parse(sn)
# B: read_excel, 
dfB= pd.read_excel(fn, sheet_name=sn)

print(dfA.columns)
