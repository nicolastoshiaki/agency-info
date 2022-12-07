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
def trata_proposta(proposta): 
    if pd.isna(proposta): 
        return 
    else: 
        proposta = proposta.replace('-', ' ') 
        proposta = proposta.replace('.', ' ') 
        proposta = proposta.replace(' ', ' ') 
        proposta = proposta.replace(' ', '') 
        proposta = proposta[:14] 
        return proposta 

def trata_protocolo(lista): 
    lista_prot = [item[0] for item in lista] 
    protocolo = lista_prot[0] 
    if protocolo is None:
        return None 
    else: 
        protocolo_ajustado = protocolo.strip() 
    return protocolo_ajustado 
        
def deleta_dados(sheet): 
    while sheet.max_row > 1: 
        sheet.delete_rows(2) 
        return 
        
# Estabelece conexão com banco de dados ---------------------------------------------------------------- 
connection_string = ("Driver={SQL Server};" "Server=SERVER;" "Database=DATABASE;" "Trusted_Coneection=yes") 
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url) 

# Declara credenciais da API Correios ------------------------------------------------------------------ 
contrato = 123456 
login = 'login_exemple' 
senha = 'password_exemplo' 

df_prop = pd.read_excel(r'path\file.xlsx') 

# Cria e Trata a Lista de Propostas -------------------------------------------------------------------- 
lista_proposta = [] 

for index, row in df_prop.iterrows():
    proposta_inserida = str(row['PROPOSTAS']) 
    proposta_ajust = trata_proposta(proposta_inserida) 
    lista_proposta.append(proposta_ajust) 

for item in lista_proposta: 
    if pd.isna(item): 
        lista_proposta.remove(item) 
    if len(lista_proposta) == 0: 
        print('Não há propostas a serem consultadas') 
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

# Cria e Trata a Lista de Protocolos ------------------------------------------------------------------ 
lista_protocolo = [] 
for item in lista_proposta:
    df = pd.read_sql('SELECT API_KEY FROM DATABSE_TABLE WITH(NOLOCK) WHERE PROPOSTA = {}'.format(item), engine) 
    lista_doc_parceiro = df.values.tolist() 
    protocolo_ajustado = trata_protocolo(lista_doc_parceiro) 
    lista_protocolo.append(protocolo_ajustado) 
    df_csv = pd.read_csv(r'path\File_with_agency_information.csv', sep=';', encoding='latin-1') 
    
lista_status_atendimento = [] 
lista_data_atendimento = [] 
lista_atendente = [] 
lista_nome_agencia = [] 
lista_mcu = [] 
lista_uf = [] 
lista_municipio = [] 
lista_endereco_agencia = [] 
lista_complem_end_agencia = [] 
lista_num_endereco = [] 
lista_bairro = [] 
lista_cep = [] 

# Consulta os protocolos na API Correios ----------------------------------------------------------------- 
for item in lista_protocolo: 
    if item is None: 
        lista_status_atendimento.append('') 
        lista_data_atendimento.append('') 
        lista_atendente.append('') 
        lista_nome_agencia.append('') 
        lista_mcu.append('') 
        lista_uf.append('') 
        lista_municipio.append('') 
        lista_endereco_agencia.append('') 
        lista_complem_end_agencia.append('') 
        lista_num_endereco.append('') 
        lista_bairro.append('') 
        lista_cep.append('') 
    else: 
        api_request = requests.get( 'api_link/contrato/{}/protocolo/{}'.format(contrato, item), auth=(login, senha)) 
        retorno = api_request.json() 
        
mcu = retorno['identificadorAgencia'].strip() 
lista_status_atendimento.append(retorno['statusAtendimento'].strip()) 
lista_data_atendimento.append(retorno['dataAtendimento'].strip()) 
lista_atendente.append(retorno['identificadorAtendente'].strip()) 
lista_nome_agencia.append(retorno['nomeAgencia'].strip()) 
lista_mcu.append(mcu) 
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
dados = list(zip(lista_proposta, lista_protocolo, lista_status_atendimento, lista_data_atendimento, lista_atendente, lista_nome_agencia, lista_mcu, lista_uf, lista_municipio, lista_endereco_agencia, lista_complem_end_agencia, lista_num_endereco, lista_bairro, lista_cep)) 

df_final = pd.DataFrame(dados, columns=['Proposta', 'Protocolo', 'Status_Atendimento', 'Data_Atendimento', 'Atendente', 'Nome_Agência', 'MCU', 'UF', 'Municipio', 'Endereco_Agência', 'Complemento_Endereco_Agência', 'Num_Endereco', 'Bairro', 'CEP']) 

# Cria a integração com o outlook 
outlook = win32.Dispatch('outlook.application') 

# Cria um email 
email = outlook.CreateItem(0) 

# Configura as informações do seu e-mail 
email.To = destinatarios 
email.Subject = "E-mail automático - Info Agências de Propostas Correios" 
email.HTMLBody = """ <p>Prezado(a),</p> Segue abaixo os dados das agências vinculadas as propostas existentes na planilha.</p> <p>{}</p> <p>Atenciosmente,</p> <p>GEGOPS - DIGOPS</p> """.format(df_final.to_html()) 
email.Send() 

# Deleta os registros da planilha ------------------------------------------------------------------------------ 
planilha = openpyxl.load_workbook(r'path\file.xlsx') 
sheet = planilha['Propostas ECT'] 
deleta_dados(sheet) 
planilha.save(r'path\file.xlsx') 
print('E-mail enviado')