# MODULO DE LEITURA DA PLANILHA DO GOOGLE

from __future__ import print_function
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pandas as pd


def boas_vindas():

    print('####################################')

    print('\nBENVINDO(A) AO DATA ESPECIAL DE COLAÇÃO DE GRAU AUTOMATIZADA\n')

    linha = int(input('Deseja realizar o registro a partir de qual linha? ')) 

    ultima_linha = int(input('Deseja realizar o registro até qual linha? ')) 


    print('####################################')

    return linha, ultima_linha




# DEFINIÇÕES GOOGLE SHEETS E INICIALIZAÇÃO DE VARIÁVEIS
# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1TWWJS_0wMSoFUUiQDbEmEROkQQPnn1M83u31t-ycGmY'
SAMPLE_RANGE_NAME = 'Data Especial de Colação de Grau!I2:X80'



class Aluno:
    def __init__(self, id, data_solicitacao, email,  nome, cpf, rg, matricula, email_alternativo, telefone, logradouro, numero_complemento, bairro, cidade, estado, cep, curso, justificativa):
        self.id = id
        self.data_solicitacao = data_solicitacao
        self.email = email
        self.nome = nome
        self.cpf = cpf
        self.rg = rg
        self.matricula = matricula
        self.email_alternativo = email_alternativo
        self.telefone = telefone
        self.logradouro = logradouro
        self.numero_complemento = numero_complemento
        self.bairro = bairro
        self.cidade = cidade
        self.estado = estado
        self.cep = cep
        self.curso = curso
        self.justificativa = justificativa



def lendo_googlesheets():

    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        
        valores = result['values']

        
        df = pd.DataFrame(valores)

        print(df.columns)


        df.to_csv('solicitacoes.csv',index=False)


        # le os dados do arquivo csv
        df = pd.read_csv('solicitacoes.csv',header=1,dtype=str)

        # elimina todas as linhas em branco
        df = df.dropna(how='all')

        # substitui todos os NaN por vazio
        df.fillna("",inplace=True)

        # insere a coluna data contendo somente a data do documento recebido 
        df.insert(0, column='data_solicitacao', value= df['Recebido em'].apply(lambda x: x.split(' ')[0]))
        
        # insere o ID no dataset 
        df.insert(0, column='id', value= range(1, 1 + len(df)))

       

        # renomeia as colunas
        df.rename(columns={'Endereço de e-mail':'email','Nome completo:':'nome','CPF:':'cpf', 'Número do documento de identidade:':'rg', 'Matrícula (se souber):':'matricula', 'E-mail:':'email_alternativo', 'Telefone, com DDD:':'telefone','Logradouro:':'logradouro','Número/Complemento: ':'numero_complemento','Bairro:':'bairro',
                   'Cidade:':'cidade','Estado:':'estado','CEP':'cep',
                   'Informe seu curso de graduação:':'curso',
                   'Escreva aqui suas justificativas e observações para o pedido, de forma clara e objetiva:':'justificativas'},inplace=True)
        
        # exclui duas colunas 
        df = df.drop(['Recebido em'], axis=1)

        # transforma a coluna ID em indice
        df.set_index('id')


        aluno = Aluno(  df.id, 
                    df.data_solicitacao,
                    df.email,
                    df.nome, 
                    df.cpf, 
                    df.rg, 
                    df.matricula, 
                    df.email_alternativo,
                    df.telefone, 
                    df.logradouro, 
                    df.numero_complemento, 
                    df.bairro, 
                    df.cidade, 
                    df.estado, 
                    df.cep, 
                    df.curso, 
                    df.justificativas
                )
    


        return(aluno)
   
    except HttpError as err:
        print(err)

aluno = lendo_googlesheets()

os.remove('solicitacoes.csv')