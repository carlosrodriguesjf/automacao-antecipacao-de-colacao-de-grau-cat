 # MODULO PARA ENVIO DE E_MAIL
from time import sleep
import pyautogui
import pyperclip

def enviar_email(email, numero_processo, numero_oficio):

    #clicando em escrever e-mail
    sleep(1)
    pyautogui.moveTo(89,208,duration=1)
    sleep(1)
    pyautogui.click()

    #clicando em destinatario
    pyautogui.click(1360,500,duration=1)
    sleep(1)

    #digitando o destinatario
    pyperclip.copy('') 
    pyautogui.write(email)
    sleep(2)
    pyautogui.hotkey('ctrl','v')
    sleep(1)
    #apertando o ENTER
    pyautogui.press('enter')
    sleep(1)

    assunto = "Solicitação de Data Especial de Colação de Grau"

    #clicando em assunto
    pyautogui.click(1337,530,duration=1)
    sleep(1)

    #digitando o assunto
    pyperclip.copy('') 
    pyperclip.copy(assunto)
    sleep(1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)
    
    

    #clicando no corpo do e-mail
    pyautogui.click(1312,558,duration=1)
    sleep(1)

    # Corpo do email
    corpo_email = f"""Prezado(a),

    Sua solicitação foi encaminhada para a análise do setor responsável.

    Nº Único de Protocolo {numero_processo}

    Documento SEI nº {numero_oficio}


    """

    #digitando o corpo do email
    pyperclip.copy('') 
    pyperclip.copy(corpo_email)
    sleep(1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    #apertando o ENTER
    pyautogui.press('enter')
    sleep(1)

    #clicando na caneta para assinatura
    pyautogui.click(1632,1009,duration=2)

    #selecionando uma assinatura
    pyautogui.click(1646, 942,duration=2)


    #clicando em enviar
    pyautogui.click(1347,1010,duration=2)
    sleep(1)
    
    print('E-mail enviado com sucesso!')
