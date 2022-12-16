import requests 
import json 
from sqlalchemy.engine import URL 
from sqlalchemy import create_engine 
import pyodbc 
import pandas as pd 
import win32com.client as win32 
import sys 
import openpyxl 

# Define as funções ---------------------------------------------------------------------- 
def trata_A(A): 
    if pd.isna(A): 
        return 
    else: 
        A = A.replace('-', ' ') 
        A = A.replace('.', ' ') 
        A = A.replace(' ', ' ') 
        A = A.replace(' ', '') 
        A = A[:14] 
        return A 

def trata_B(lista): 
    lista_prot = [item[0] for item in lista] 
    B = lista_prot[0] 
    if B is None:
        return None 
    else: 
        B_ajustado = B.strip() 
    return B_ajustado 
        
def deleta_dados(sheet): 
    while sheet.max_row > 1: 
        sheet.delete_rows(2) 
        return 
        
# Estabelece conexão com banco de dados ---------------------------------------------------------------- 
connection_string = ("Driver={SQL Server};" "Server=SERVER;" "Database=DATABASE;" "Trusted_Coneection=yes") 
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url) 

# Declara credenciais da API ------------------------------------------------------------------ 
contrato = 123456 
login = 'login_exemple' 
senha = 'password_exemple' 

df_prop = pd.read_excel(r'path\file.xlsx') 

# Cria e Trata a Lista de A -------------------------------------------------------------------- 
lista_A = [] 

for index, row in df_prop.iterrows():
    A_inserida = str(row['A']) 
    A_ajust = trata_A(A_inserida) 
    lista_A.append(A_ajust) 

lista_A2 = [item for item in lista_A if str(item) != 'nan'] 
        
if len(lista_A2) == 0: 
    print('Não há A a serem consultadas') 
    sys.exit() 
        

        
# Cria e Trata a Lista de E-mails --------------------------------------------------------------------- 
lista_email = [] 
for index, row in df_prop.iterrows(): 
    email_inserido = str(row['EMAILS']) 
    lista_email.append(email_inserido) 
    
lista_email2 = [item for item in lista_email if str(item) != 'nan'] 

if len(lista_email2) == 0: 
    print('Não há e-mail inserido na planilha') 
    sys.exit() 
        
destinatarios = '; '.join(lista_email2) 

# Cria e Trata a Lista de B ------------------------------------------------------------------ 
lista_B = [] 
for item in lista_A2:
    df = pd.read_sql('SELECT API_KEY FROM DATABSE_TABLE WITH(NOLOCK) WHERE A = {}'.format(item), engine) 
    lista_doc = df.values.tolist() 
    B_ajustado = trata_B(lista_doc) 
    lista_B.append(B_ajustado) 

df_csv = pd.read_csv(r'path\File_with_agency_information.csv', sep=';', encoding='latin-1') 
    
lista_status_atendimento = [] 
lista_data_atendimento = [] 
lista_atendente = [] 
lista_nome_agencia = [] 
lista_C = [] 
lista_uf = [] 
lista_municipio = [] 
lista_endereco_agencia = [] 
lista_complem_end_agencia = [] 
lista_num_endereco = [] 
lista_bairro = [] 
lista_cep = [] 

# Consulta os B na API ----------------------------------------------------------------- 
for item in lista_B: 
    if item is None: 
        lista_status_atendimento.append('') 
        lista_data_atendimento.append('') 
        lista_atendente.append('') 
        lista_nome_agencia.append('') 
        lista_C.append('') 
        lista_uf.append('') 
        lista_municipio.append('') 
        lista_endereco_agencia.append('') 
        lista_complem_end_agencia.append('') 
        lista_num_endereco.append('') 
        lista_bairro.append('') 
        lista_cep.append('') 
    else: 
        api_request = requests.get('api_link/contrato/{}/B/{}'.format(contrato, item), auth=(login, senha)) 
        retorno = api_request.json() 
        
        mcu = retorno['identificadorAgencia'].strip() 
        lista_status_atendimento.append(retorno['statusAtendimento'].strip()) 
        lista_data_atendimento.append(retorno['dataAtendimento'].strip()) 
        lista_atendente.append(retorno['identificadorAtendente'].strip()) 
        lista_nome_agencia.append(retorno['nomeAgencia'].strip()) 
        lista_C.append(mcu) 
        lista_uf.append(retorno['uf'].strip()) 
        lista_municipio.append(retorno['municipio'].strip()) 
        endereco_agencia = df_csv.loc[df_csv['MCU'] == int(mcu), 'ENDERECO'].iloc[0] 
        complem_end_agencia = df_csv.loc[df_csv['MCU'] == int(mcu), 'COMPL_ENDERECO'].iloc[0] 
        num_endereco = df_csv.loc[df_csv['MCU'] == int(mcu), 'Número'].iloc[0] 
        bairro = df_csv.loc[df_csv['MCU'] == int(mcu), 'BAIRRO'].iloc[0] 
        cep = df_csv.loc[df_csv['MCU'] == int(mcu), 'CEP'].iloc[0] 

        if pd.isna(endereco_agencia): 
            endereco_agencia = '' 
        if pd.isna(complem_end_agencia): 
            complem_end_agencia = '' 
        if pd.isna(num_endereco): 
            num_endereco = '' 
        if pd.isna(bairro): 
            bairro = '' 
        if pd.isna(cep): 
            cep = '' 
    
        lista_endereco_agencia.append(endereco_agencia) 
        lista_complem_end_agencia.append(complem_end_agencia) 
        lista_num_endereco.append(num_endereco) 
        lista_bairro.append(bairro) 
        lista_cep.append(cep) 

# Monta o dataframe final que vai no corpo do e-mail ---------------------------------------------------------- 
dados = list(zip(lista_A2, lista_B, lista_status_atendimento, lista_data_atendimento, lista_atendente, lista_nome_agencia, lista_C, lista_uf, lista_municipio, lista_endereco_agencia, lista_complem_end_agencia, lista_num_endereco, lista_bairro, lista_cep)) 

df_final = pd.DataFrame(dados, columns=['A', 'B', 'Status_Atendimento', 'Data_Atendimento', 'Atendente', 'Nome_Agência', 'C', 'UF', 'Municipio', 'Endereco_Agência', 'Complemento_Endereco_Agência', 'Num_Endereco', 'Bairro', 'CEP']) 

# Cria a integração com o outlook 
outlook = win32.Dispatch('outlook.application') 

# Cria um email 
email = outlook.CreateItem(0) 

# Configura as informações do seu e-mail 
email.To = destinatarios 
email.Subject = "E-mail automático - Info Agências de A" 
email.HTMLBody = """ <p>Prezado(a),</p> Segue abaixo os dados das agências vinculadas as A existentes na planilha.</p> <p>{}</p> <p>Atenciosamente,</p> <p>GERENCIA - DIRETORIA</p> """.format(df_final.to_html()) 
email.Send() 

# Deleta os registros da planilha ------------------------------------------------------------------------------ 
planilha = openpyxl.load_workbook(r'path\file.xlsx') 
sheet = planilha['Sheet_name'] 
deleta_dados(sheet) 
planilha.save(r'path\file.xlsx') 
print('E-mail enviado')
