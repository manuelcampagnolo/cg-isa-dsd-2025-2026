######################################################
# Manuel Campagnolo (mar 2024)
# ISA/ULIsboa
# script para preparar ficheiro Excel DSD ISA 2024-2025
#######################################################

from pathlib import Path
import pandas as pd
from copy import copy
from unidecode import unidecode
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.worksheet.protection import SheetProtection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import quote_sheetname # https://openpyxl.readthedocs.io/en/latest/validation.html?highlight=validation
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter, column_index_from_string, coordinate_to_tuple
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors
from functions import simplify_strings, df_to_excel_with_columns, compact_excel_file, get_letter_from_column_name,unlock_cells,stripe_cells
from functions import add_suffix_to_duplicates, reorder_and_filter_dataframe, insert_row_at_beginning, insert_row_at_end,sort_list
from functions import replace_values_in_string
import warnings
warnings.filterwarnings("ignore", category=UserWarning)
############################################# constants

UNPROTECT_OUTPUT_CELLS=True # para desproteger células com função unlock_cells(ws, col_name, min_row=None, max_row=None):
PROTECT_WORKSHEET=True # tem que ser True para impedir escrita
PASSWORD='kathleen'

# root (working directory)
try:
    working_dir=Path(__file__).parent.parent # working directory from script location: scripts are in 'scripts' folder
except:
    working_dir=Path().absolute()


# folders
FOLDER_SERVICOS= 'ficheiros_servicos_ISA'
FOLDER_OUTPUT='output_files'
FOLDER_FICH_RESPONSAVEIS_UCs='ficheiros_responsaveis_ucs'
FOLDER_2425='DSD_2024_2025'
FOLDER_2324='DSD_2023_2024'
FN_responsaveis='DSD_2024_2025_responsaveis_UCs_fechado_12_mar_2024_desprotegido_ML_15MAR2024_corr_SIG2ciclo.xlsx'
FN_resumo='resumo_DSD_2324_feito_em_fev_2024.xlsx'

# list files in FOLDER_OUTPUT and create name of output file
folder_output=working_dir/FOLDER_2425/FOLDER_OUTPUT
number_files=len(list(folder_output.rglob("DSD_2024_2025*")))
DSD_OUTPUT_FICH='DSD_2024_2025_v'+str(number_files+1)+'.xlsx' 

# main files
fn_resp= working_dir / FOLDER_2425 / FOLDER_FICH_RESPONSAVEIS_UCs / FN_responsaveis
fn_resumo= working_dir / FOLDER_2324 / FN_resumo
fn_output = folder_output / DSD_OUTPUT_FICH

# worksheets (upper case) and column names (lower case) from fn_resp
UC ='uc_2024-25'
UC_uc = 'unidade_curricular'
UC_obrigatoria = 'uc_obrigatoria'
UC_optativa = 'uc_optativa'
UC_codigo='cod_uc'
UC_resp='responsavel_unidade_curricular'
UC_sugestoes='sugestões de modificação da info da UC'
UC_autor_sugestao='autor da sugestão'
UC_ciclo_curso = 'ciclo_curso'
UC_horas_contacto = 'h_contacto_total'
UC_ano='ano'
UC_sem='sem'
UC_horas_totais_sugeridas='horas totais sugeridas'
UC_horas_justif='justificação horas totais'
UCMETA='uc_meta'
RH='RH'
RH_nome='nome' # nomes docentes
N_extra_nomes=0
RH_numero = 'num_pessoal'
RH_data_fim='data_fim'
RH_posicao='posicao'
RH_obs = 'Obs'
RHPOSICAO='RH_posicao' # drop
RHMETA='RH_meta'
PE='planos_estudos' # drop
AC='AC' #drop
ACMETA='AC_meta'#drop
POS='POS' #drop
UCAREA='uc_AreaCient' # drop
CODCURSO='cod_curso'
CODCURSO_cod_curso='cod_curso'
CODCURSO_sigla='siglaCurso'

# sheet from fn_resumo
HORASUCS='horas_UCs'
HORASUCS_codigo='Código UC'
HORASUCS_nome_uc='Nome da UC'
HORASUCS_total_horas='Total horas UC'
HORASUCS_total_horas_novo = 'Total_DSD_2324'

# values: Sheet_column_value
RH_nome_pro_bono='docente_PRO_BONO' 
RH_nome_em_contratacao='Docente em contratação'
RH_data_fim_sem_termo='sem termo'
DATA_TERMO_CERTO='2024-09-01'

# desproteger células:
# responsável
# docentes extra + observações

'''
, 'UC_meta': 'uc_meta', 'AC': 'uc_AreaCient', 'AC_meta': 'AC_meta', 'CC': 'cod_curso', 'RH': 'RH', 'RH_meta': 'RH_meta', 'POS': 'RH_posicao'}
UC = {'UC_ciclo_': 'ciclo_curso', 'UC_uc_obr': 'uc_obrigatoria', 'UC_uc_opt': 'uc_optativa', 'UC_cod_uc': 'cod_uc', 'UC_unidad': 'unidade_curricular', 'UC_Respon': 'Responsavel', 'UC_não_fu': 'não_funciona_2024_2025', 'UC_area_c': 'area_cient', 'UC_ano': 'ano', 'UC_sem': 'sem', 'UC_h_trab': 'h_trab_totais', 'UC_T': 'T', 'UC_TP': 'TP', 'UC_PL': 'PL', 'UC_TC': 'TC', 'UC_S': 'S', 'UC_E': 'E', 'UC_OT': 'OT', 'UC_O': 'O', 'UC_h_cont': 'h_contacto_total', 'UC_ECTS': 'ECTS', 'UC_uc_int': 'uc_interna', 'UC_obs': 'obs'}
UC_meta = {'UC_meta_ciclo_': 'ciclo_curso', 'UC_meta_1 cicl': '1 ciclo', 'UC_meta_Licenc': 'Licenciatura'}
AC = {'AC_area_c': 'area_cientifica', 'AC_sigla': 'sigla', 'AC_classi': 'classificacao_area'}
AC_meta = {'AC_meta_classi': 'classificacao_area', 'AC_meta_CNAEF': 'CNAEF', 'AC_meta_Classi': 'Classificação Nacional das Áreas de Educação e Formação (CNAEF), Portaria n.o 256/2005 de 16 de Março. Classificação: nível 1, 2 e 3 (código: 3 dígitos)'}
CC = {'CC_cod_cu': 'cod_curso', 'CC_grau': 'grau', 'CC_curso': 'curso', 'CC_obs': 'obs'}
RH = {'RH_num_pe': 'num_pessoal', 'RH_nome': 'nome', 'RH_corpo': 'corpo', 'RH_posica': 'posicao', 'RH_ETI': 'ETI', 'RH_contra': 'contratacao_vinculo', 'RH_data_f': 'data_fim', 'RH_Obs': 'Obs'}
RH_meta = {'RH_meta_num_pe': 'num_pessoal', 'RH_meta_número': 'número de indentificação pessoal atribuído pelos RH', 'RH_meta_Unname': 'Unnamed: 4'}
POS = {'POS_posica': 'posicao', 'POS_h_min': 'h_min', 'POS_h_max': 'h_max', 'POS_obs': 'obs'}
''' 

# cor light red, light yellow
fill_red = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")
fill_green = PatternFill(start_color="C0FFCB", end_color="C0FFCB", fill_type="solid")
fill_yellow = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid") # alpha: 1st 2 characters
fill_light_yellow = PatternFill(start_color="FFFFED", end_color="FFFFED", fill_type="solid") 
thin_border=Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

###################################################################### read inputs to dataframes
# read sheet names
sheets_resp = pd.ExcelFile(fn_resp).sheet_names
sheets_resumo = pd.ExcelFile(fn_resumo).sheet_names

# read data from inputs
df_rh = pd.read_excel(fn_resp, sheet_name=RH)
df_rh_meta = pd.read_excel(fn_resp, sheet_name=RHMETA,header=None)
df_rh_meta.columns=['coluna','descricao_1','descricao_2','descricao_3','descricao_4']
df_uc = pd.read_excel(fn_resp, sheet_name=UC)
df_uc[UC_codigo]=df_uc[UC_codigo].astype(str).str.strip()
df_uc_meta = pd.read_excel(fn_resp, sheet_name=UCMETA,header=None)
df_uc_meta.columns=['coluna','descricao','observacoes']
df_cod_curso = pd.read_excel(fn_resp, sheet_name=CODCURSO)
df_cod_curso[CODCURSO_cod_curso]=df_cod_curso[CODCURSO_cod_curso].astype(str).str.strip()
df_resumo = pd.read_excel(fn_resumo, sheet_name=HORASUCS)
df_resumo[HORASUCS_codigo]=df_resumo[HORASUCS_codigo].astype(str).str.strip()

########################################################################## process dataframes

# main dataframe for DSD
# select columns to be shown
df=df_uc[[UC_ciclo_curso,UC_obrigatoria,UC_optativa,UC_codigo,UC_resp,UC_uc, UC_horas_contacto,UC_ano, UC_sem, UC_autor_sugestao, UC_sugestoes]]
# merge with df_resumo: para ir buscar o total de horas da UC (2023/2024)
df=pd.merge(df,df_resumo[[HORASUCS_codigo,HORASUCS_nome_uc,HORASUCS_total_horas]],left_on=UC_codigo,right_on=HORASUCS_codigo,how='left')
df=df.fillna(' ')
# adicionar colunas para horas revistas e justificação da mudança
df[UC_horas_totais_sugeridas]= df[HORASUCS_total_horas]
df[UC_horas_justif]=' '
# trocar abreviaturas em UC_obrigatoria,UC_optativa
df_replace=df_cod_curso[[CODCURSO_cod_curso,CODCURSO_sigla]]
df[UC_obrigatoria] = df[UC_obrigatoria].apply(lambda x: replace_values_in_string(x, df_replace))
df[UC_optativa] = df[UC_optativa].apply(lambda x: replace_values_in_string(x, df_replace))

# RH: acrescentar coluna horas totais do docente
df_rh['horas']

print(df[[UC_obrigatoria,UC_optativa]][:15])

# select columns for printing
dfaux=df[[UC_ciclo_curso,UC_codigo,UC_uc, UC_horas_contacto,HORASUCS_codigo,HORASUCS_total_horas, UC_horas_totais_sugeridas, UC_horas_justif]]
#print(dfaux[dfaux[HORASUCS_total_horas].notna()])

# Criar novo registo em _meta
df = df.rename(columns={HORASUCS_total_horas: HORASUCS_total_horas_novo})
df_uc_meta=insert_row_at_end(df_uc_meta, {'coluna':HORASUCS_total_horas_novo, 'descricao':'horas totais de docência da UC no DSD 2023-2024'})


