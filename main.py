# MODULO PRINCIPAL

#importações
import processo_sei
import leitura_googlesheets
from envio_email import enviar_email
from  gravacao_googlesheets import gravando_google_sheets



if __name__ == '__main__':

    linha, ultima_linha =  leitura_googlesheets.boas_vindas()

    for linha in range(linha, ultima_linha + 1):

        aluno = leitura_googlesheets.lendo_googlesheets()

        email, numero_processo, numero_oficio =  processo_sei.preenchendo_processo_sei(aluno, linha, ultima_linha)

        enviar_email(email, numero_processo, numero_oficio)

        gravando_google_sheets(linha, ultima_linha, numero_processo, numero_oficio)

            
print('OPERAÇÃO FINALIZADA')
