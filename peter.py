######################################################
# Manuel Campagnolo (abril 2023)
# ISA/ULIsboa
# script para preparar ficheiro Excel DSD ISA 2023-2024
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

import pandas as pd
from datetime import datetime
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

########################################################################################################
fnDSD="DSD_2324_2maio2023.xlsx" #  "DSD_v1_teste.xlsx"; 
fnResumo="resumo_DSD_2324_2maio.xlsx"
VALIDATION_VALUE='Inserir docente'
# worksheets de input
ws_name_preencher='DSD (para preencher)'
ws_name_info='DSD (informação UCs)'
ws_name_docentes='DocentesNovo' 
# ws preencher

coldoc='AK'
print(column_index_from_string(coldoc))
colunas_red=['Grandes Áreas Científicas (FOS)', 'Áreas Científicas (FOS)', 'Áreas Disicplinares', 'Departamento',  'Responsável', 'Nome da UC', 'ciclo de estudos', 'Código UC', 'ano curricular', 'semestre', 'ECTS', 'semanas', 'Total horas da UC', 'Somatório', 'Horas em falta na UC']
column_codigo_UC_preencher='Código UC'
column_horas_em_falta_preencher='Horas em falta na UC'
column_somatorio_preencher= 'Somatório'
column_preencher_first='Total Horas Teóricas'
column_preencher_last='Total Horas  Outras'
column_new_horas_totais='Horas totais docente'
column_new_horas_semanais='Horas semanais docente'
info_cols_to_unprotect=['H','M','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AD','AE']
DAA='Docente a atribuir'
# ws info (colocada para facilitar VLOOKUP)
column_total_horas_info='Total Horas previsto '
column_codigo_UC_info='Código UC'
# colunas DSD
columns_to_copy='Nome da UC' # 'Nome' # designação UC
column_key='Inserir docentes na UC ' #'docentes ' # cuidado: tem um espaço a mais
coluna_validacao=column_key
col_codigo_uc='Código UC'
column_horas_em_falta_preencher='Horas em falta na UC' #Horas em falta na UC
column_somatorio_preencher= 'Somatório'
colunas_nome_UC=['Nome da UC', column_codigo_UC_info, 'ciclo de estudos', 'ano curricular', 'semestre'] #Nome da UC
colunas_FOS=['Grandes Áreas Científicas (FOS)', 'Áreas Científicas (FOS)', 'Áreas Disicplinares', 'Departamento']
colResponsavel='Responsável'
newcolResponsavel='nomeResponsavel'
colNomeDocente='NomeDocente'
colLinhaResp='linhaResp'
cols_horas=['Total Horas Teóricas', 'Horas Teóricas por semana','Total Horas Teórico-práticas', 'Horas Teórico-práticas por semana', 'Total Horas Laboratoriais', 'Horas Laboratoriais por semana', 'Total Horas  Trabalho de campo', 'Horas Trabalho de Campo por semana','Total Horas  Seminário', 'Horas Seminário por semana','Total Horas  Estágio', 'Total Horas  Outras']
# colunas DSD docentes
col_horas_semanais='Horas semanais docente'
col_horas_totais='Horas totais docente'
col_nome_completo='Nome completo'
col_posicao='Posição'
cols_horas_docentes=[col_horas_semanais,col_horas_totais,col_nome_completo,col_posicao]

# colunas a apagar em DSD info: funciona
col_apagar_info_1='Total Horas Somadas '
col_apagar_info_2='Horas em falta na UC'
col_apagar_info_3='Código UC'
col_apagar_info_4='Total Horas previsto ' #Total Horas previsto 
################################################################################################################# funcoes
def fill_empty_cells(df, col_name):
    # Create a copy of the dataframe to avoid modifying the original
    filled_df = df.copy()
    
    # Create a mask of empty cells in the column
    empty_mask = filled_df[col_name].isna()
    
    # Iterate over each row in the column
    for i, cell_value in enumerate(filled_df[col_name]):
        # If the cell is empty, fill it with the value of the closest non-empty cell above
        if pd.isna(cell_value):
            j = i - 1
            while j >= 0 and pd.isna(filled_df.loc[j, col_name]):
                j -= 1
            if j >= 0:
                filled_df.loc[i, col_name] = filled_df.loc[j, col_name]           
    return filled_df

# def df_to_excel_simple(df, ws, header=True, index=True):
#     for r in dataframe_to_rows(df, index=index, header=header):
#         ws.append(r)

def df_to_excel(df, ws, header=True, index=True, startrow=0, startcol=0):
    """Write DataFrame df to openpyxl worksheet ws"""
    rows = dataframe_to_rows(df, header=header, index=index)
    for r_idx, row in enumerate(rows, startrow + 1):
        for c_idx, value in enumerate(row, startcol + 1):
             ws.cell(row=r_idx, column=c_idx).value = value

def create_workbook_from_dataframe(df):
    """
    1. Create workbook from specified pandas.DataFrame
    2. Adjust columns width to fit the text inside
    3. Make the index column and the header row bold
    4. Fill background color for the header row

    Other beautification MUST be done by usage side.
    """
    workbook = openpyxl.Workbook()
    ws = workbook.active

    rows = dataframe_to_rows(df.reset_index(), index=False)
    col_widths = [0] * (len(df.columns) + 1)
    for i, row in enumerate(rows, 1):
        for j, val in enumerate(row, 1):

            if type(val) is str:
                cell = ws.cell(row=i, column=j, value=val)
                col_widths[j - 1] = max([col_widths[j - 1], len(str(val))])
            elif hasattr(val, "sort"):
                cell = ws.cell(row=i, column=j, value=", ".join(list(map(lambda v: str(v), list(val)))))
                col_widths[j - 1] = max([col_widths[j - 1], len(str(val))])
            else:
                cell = ws.cell(row=i, column=j, value=val)
                col_widths[j - 1] = max([col_widths[j - 1], len(str(val)) + 1])

            # Make the index column and the header row bold
            if i == 1 or j == 1:
                cell.font = Font(bold=True)

    # Adjust column width
    for i, w in enumerate(col_widths):
        letter = get_column_letter(i + 1)
        ws.column_dimensions[letter].width = w

    return workbook

def set_border(ws):
    border = Border(left=Side(border_style='thin', color='000000'),
                right=Side(border_style='thin', color='000000'),
                top=Side(border_style='thin', color='000000'),
                bottom=Side(border_style='thin', color='000000'))

    rows = ws.iter_rows()
    for row in rows:
        for cell in row:
            cell.border = border

############################################################ funcoes
# devolve letra da coluna com nome (1a linha) da worksheet ws
def nomeColuna2letter(ws,nome):
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx=headers.index(nome)
    return get_column_letter(idx+1)

####################################################################################################################

dfinfo= pd.read_excel(fnDSD, sheet_name=ws_name_info) 
dfdsd= pd.read_excel(fnDSD, sheet_name=ws_name_preencher)
dfdsd.columns

# limpar dfinfo
dfinfo=dfinfo.drop(col_apagar_info_1,axis=1)
dfinfo=dfinfo.drop(col_apagar_info_2,axis=1)
dfinfo=dfinfo.drop(col_apagar_info_3,axis=1)
dfinfo=dfinfo.drop(col_apagar_info_4,axis=1)
dfinfo = dfinfo.dropna(axis=1, how='all')

#selecionar colunas dos docentes em DSD
dfhd=dfdsd[cols_horas_docentes]
dfhd=dfhd.dropna(axis=0)
dfhd[col_horas_semanais]=dfhd[col_horas_semanais].round(2)

# criar coluna de Responsável
dfdsd.loc[dfdsd[colResponsavel]=='sim',newcolResponsavel]=dfdsd.loc[dfdsd[colResponsavel]=='sim',column_key]
#dfdsd.head(20)[newcolResponsavel]

# preencher algumas colunas com o valor acima
for col in colunas_FOS+colunas_nome_UC+[newcolResponsavel]:
    dfdsd=fill_empty_cells(dfdsd, col)
print(dfdsd.shape)

# remover linhas com VALIDATION_VALUE ou em branco na coluna column_key
dfdsd=dfdsd[dfdsd[column_key]!= VALIDATION_VALUE]
dfdsd=dfdsd[dfdsd[column_key].notna()]
print(dfdsd.shape)

# drop 2nd row
#dfdsd=dfdsd.drop(1)

# remover colunas DSD
dfdsd=dfdsd[colunas_FOS+[colResponsavel, newcolResponsavel,column_key]+colunas_nome_UC+cols_horas+[column_horas_em_falta_preencher]]

# selecionar as linhas dos responsáveis das UCs
dfucs=dfdsd[dfdsd[colResponsavel]=='sim']
dfucs=dfucs[colunas_FOS+[newcolResponsavel]+colunas_nome_UC+[column_horas_em_falta_preencher]]

# eliminar as linhas dos responsáveis
dfdsd=dfdsd[dfdsd[colResponsavel]!='sim']

############### output
wbr=openpyxl.Workbook()
wsr_ucs=wbr.create_sheet('horas_UCs')
wsr_docentes=wbr.create_sheet('horas_docentes')
wsr_ucs_docentes=wbr.create_sheet('horas_UCs_docentes')
wsr_info=wbr.create_sheet('info_UCs')

def df_to_excel_with_columns(df,ws,maxwidth=30,header=True,index=False):
    for column in df.columns:
        # get the column letter
        column_letter = get_column_letter(df.columns.get_loc(column) + 1)
        # determine the optimal width based on the contents of the column
        max_length = df[column].astype(str).map(len).max()
        width = min(max_length+2, maxwidth) # set a maximum width of 30
        # set the column width
        ws.column_dimensions[column_letter].width = width
        # write 
        df_to_excel(df,ws,header=True,index=False)

df_to_excel_with_columns(dfucs, wsr_ucs)
df_to_excel_with_columns(dfdsd, wsr_ucs_docentes)
df_to_excel_with_columns(dfhd, wsr_docentes)
df_to_excel_with_columns(dfinfo, wsr_info)

if 'Sheet' in wbr.sheetnames:  # remove default sheet
    wbr.remove(wbr['Sheet'])

wbr.save(fnResumo)
wbr.close


