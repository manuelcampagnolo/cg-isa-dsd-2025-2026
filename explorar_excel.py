######################################################
# Manuel Campagnolo (abril 2023)
# ISA/ULIsboa
# script para explorar um Excel
#######################################################

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter, column_index_from_string, coordinate_to_tuple
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.styles import Protection
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors
from openpyxl.styles import Border, Side
from copy import copy
import os
from fuzzywuzzy import fuzz # compare strings
import numpy as np
from fuzzywuzzy import process
import pandas as pd
from datetime import datetime
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

# Tabela Recursos Humanos
fn='nomes_docentes_codigos_RH_17maio2023_editado_MLC.xlsx'
df= pd.read_excel(fn) 


for col in df.columns:
    print(col)
    if len(df[col].unique())<50:
        print(df[col].unique())

# Tabela Div Académica
fn='DSD_2324_12jun2023_CorrigidoML_desprotegido_editado_MLC_29junho.xlsx'
df= pd.read_excel(fn, sheet_name='DSD (informação UCs)')

for col in df.columns:
    print(col)
    if len(df[col].unique())<50:
        print(df[col].unique())

# Tabela docemtes externos # 
fn='DSD_2023_2024_servico_externo_v6_revisto_TF_DSD_28junho.xlsx'
df= pd.read_excel(fn)

count=0
for col in df.columns:
    count+=1
    if count>40: break
    print(col)
    if len(df[col].unique())<20:
        print(df[col].unique())
