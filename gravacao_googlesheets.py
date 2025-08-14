# MODULO PARA GRAVAÇÃO NA PLANILHA DO GOOGLE


from __future__ import print_function
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime
from processo_sei import numero_processo, numero_oficio



# DEFINIÇÕES GOOGLE SHEETS E INICIALIZAÇÃO DE VARIÁVEIS
# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1TWWJS_0wMSoFUUiQDbEmEROkQQPnn1M83u31t-ycGmY'
SAMPLE_RANGE_NAME = 'Data Especial de Colação de Grau!B3:H80'

ATRIBUIDO = 'CARLOS'
STATUS = 'PROTOCOLADO'
OBSERVACAO = ''
SETOR = 'PROGRAD'    
DATA_ATUAL = datetime.date.today().strftime("%d/%m/%Y")



def gravando_google_sheets(linha, ultima_linha, numero_processo, numero_oficio):

    def update_values(spreadsheet_id, range_name, value_input_option, linha, ultima_linha, numero_processo, numero_oficio,
                    _values):

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
            values = [
                [ATRIBUIDO, STATUS, OBSERVACAO, SETOR, numero_processo, numero_oficio, DATA_ATUAL],
        
            ]
            body = {
                'values': values
            }
            result = service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id, range=range_name,
                valueInputOption=value_input_option, body=body).execute()
        
            return result
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error
        
    primeira_linha = linha + 3
    # ultima = total_linhas + 3
    
    for i in range(primeira_linha, primeira_linha + 1):
        RANGE_NAME = f'Data Especial de Colação de Grau!B{i}:H{i}'

        update_values(SAMPLE_SPREADSHEET_ID,
                    RANGE_NAME, "USER_ENTERED",
                    [
                        ['A', 'B'],
                        ['C', 'D']
                        ])
        

