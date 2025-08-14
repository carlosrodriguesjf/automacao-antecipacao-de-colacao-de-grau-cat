
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from tratamento_cpf import tratar_cpf
from leitura_googlesheets import aluno
from decouple import config
import pyperclip
import pyautogui
import datetime
import os




def preenchendo_processo_sei(aluno, linha, ultima_linha):
    def iniciar_driver():
        chrome_options = Options()
        arguments = ['--lang=pt-BR', '--start-maximized', '--incognito', '--disable-extensions']
        for argument in arguments:
            chrome_options.add_argument(argument)

            chrome_options.add_experimental_option('prefs', {
            'download.prompt_for_download': False,
            'profile.default_content_setting_values.notifications': 2,
            'profile.default_content_setting_values.automatic_downloads': 1,

        })
        driver = webdriver.Chrome(service=ChromeService(
            ChromeDriverManager().install()), options=chrome_options)

        return driver

    driver = iniciar_driver()

    driver.get('https://sei.ufjf.br/') 


    # clicando login acesso SEI 
    sleep(3)
    campo_usuario_sei = driver.find_element(By.XPATH,"//div/input[@id='txtUsuario']") 
    # campo_usuario_sei.send_keys(operador.cpf_operador)
    campo_usuario_sei.send_keys(config('CPF_OPERADOR'))

    # clicando senha acesso SEI
    sleep(1)
    campo_senha_sei = driver.find_element(By.XPATH,"//div/input[@id='pwdSenha']") 
    # campo_senha_sei.send_keys(operador.senha_operador)
    campo_senha_sei.send_keys(config('SENHA_OPERADOR'))

    # clicando botao Acessar
    sleep(1)
    botao_login_sei = driver.find_element(By.XPATH,"//div/button[@id='sbmLogin']") 
    botao_login_sei.click()


    linha = linha - 3



    # salvando a janela atual
    janela = driver.current_window_handle


    # criando processo:
    sleep(3)
    botao_iniciar_processo = driver.find_element(By.XPATH,"//div//ul/li[3]/a[contains(text(), 'Iniciar Processo')]") 
    botao_iniciar_processo.click()

    # escolhendo o tipo de processo GERAL 01: Ofício (para solicitações, informações e encaminhamentos)
    sleep(1)
    botao_tipo_processo = driver.find_element(By.XPATH,"//table/tbody//tr[3]/td/a[2]") 
    botao_tipo_processo.click()


    # digitando a especificação do processo
    sleep(1)
    campo_especificacao_sei = driver.find_element(By.XPATH,"//div/input[@id='txtDescricao']") 
    # campo_especificacao_sei.send_keys(f'Data Especial de Colação de Grau - {aluno.nome[linha]}')
    campo_especificacao_sei.send_keys(f'Data Especial de Colação de Grau - {aluno.nome[linha]}')


    # definindo o nível de acesso
    sleep(1)
    botao_nivel_acesso_sei = driver.find_element(By.XPATH,"//div/input[@id='optRestrito']") 
    botao_nivel_acesso_sei.click()

    # definindo a hipótese
    sleep(1)
    botao_nivel_acesso_sei = driver.find_element(By.XPATH,"//div[@id='divNivelAcesso']/fieldset/select[@id='selHipoteseLegal']") 
    botao_nivel_acesso_sei.click()
    sleep(1)
    botao_nivel_acesso_sei.send_keys('Informação')
    sleep(1)
    botao_nivel_acesso_sei.send_keys(Keys.ENTER)
    sleep(1)

    # salvando o processo criado
    sleep(1)
    botao_salvar_processo_sei = driver.find_element(By.XPATH,"//div/div[@id='divInfraBarraComandosInferior']/button[@id='btnSalvar']") 
    botao_salvar_processo_sei.click()


    # CRIANDO DOCUMENTO

    sleep(1)
    iframe = driver.find_element(By.XPATH,"//div[@id='divConteudo']//iframe[@id='ifrVisualizacao']")
    driver.switch_to.frame(iframe)


    # clicando em novo documento
    sleep(1)
    botao_incluir_doc = driver.find_element(By.XPATH,"//div[@id='divArvoreAcoes']/a[1]")
    botao_incluir_doc.click()

    sleep(1)
    campo_tipo_doc = driver.find_element(By.XPATH,"//div/input[@id='txtFiltro']")
    campo_tipo_doc.send_keys('GERAL 01: Oficio')

    sleep(1)
    opcao_oficio = driver.find_element(By.XPATH,"//table/tbody/tr[@data-desc='geral 01: oficio']/td//a[2]/span[@class='infraSpanRealce']")
    opcao_oficio.click()

    # escolhendo o texto padrão
    botao_texto_padrao = driver.find_element(By.XPATH,"//fieldset[@id='fldTextoInicial']//div[2][@id='divOptTextoPadrao']")
    botao_texto_padrao.click()

    # escolhendo o texto padrão
    campo_texto_padrao = driver.find_element(By.XPATH,"//fieldset[@id='fldTextoInicial']/select[@id='selTextoPadrao']")
    campo_texto_padrao.click()
    campo_texto_padrao.send_keys('Data Especial de Colação de Grau')
    campo_texto_padrao.send_keys(Keys.ENTER)

    # escolhendo o nível de acesso do documento
    sleep(1)
    radio_restrito = driver.find_element(By.ID,"optRestrito")
    radio_restrito.click()


    # definindo a hipótese legal
    sleep(1)
    botao_nivel_acesso_sei = driver.find_element(By.XPATH,"//div[@id='divNivelAcesso']/fieldset/select[@id='selHipoteseLegal']") 
    botao_nivel_acesso_sei.click()
    sleep(1)
    botao_nivel_acesso_sei.send_keys('Informação')
    sleep(1)
    botao_nivel_acesso_sei.send_keys(Keys.ENTER)
    sleep(1)

    # clicando  em confirmar
    sleep(1)
    botao_confirmar =  driver.find_element(By.XPATH,"//div//div[@id='divInfraBarraComandosInferior']/button[@id='btnSalvar']")
    botao_confirmar.click()
    sleep(2)


    #PEGANDO NUMERO DO PROCESSO

    driver.switch_to.window(driver.window_handles[0])
    sleep(2)
    # salvando a janela atual
    janela_lateral = driver.current_window_handle

    sleep(1)
    iframe = driver.find_element(By.XPATH,"//div[@id='divConteudo']/iframe[1]")
    driver.switch_to.frame(iframe)

    # clicando no numero do processo
    sleep(1)
    botao_numero_processo = driver.find_element(By.XPATH,"//div[@id='header']/div/a[@class='clipboard']")
    botao_numero_processo.click()
    sleep(1)

    # pegando o numero do processo
    numero_processo = pyperclip.paste()


    sleep(1)
    driver.switch_to.window(driver.window_handles[1])

    #Maximizando a janela
    driver.maximize_window()


    #### TRATAMENTO CPF  ####
    cpf_tratado = tratar_cpf(aluno.cpf[linha])



    # Digitando o assunto no ofício
    sleep(1)
    # driver.execute_script("window.scrollTo(0, 500);")
    iframe = driver.find_element(By.XPATH,"//div[@id='divEditores']//div[3]/div/div/iframe")
    driver.switch_to.frame(iframe)


    # Localize a célula e seu texto atual nome da celula
    celula = driver.find_element(By.XPATH, "/html/body/p[2]/strong[2]")
    texto_substituir =  f"Data especial de Colação de Grau - {aluno.nome[linha]}  - {cpf_tratado}"
    # texto_substituir =  f"Data especial de Colação de Grau - {aluno.nome[linha]}"
    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, texto_substituir)


    # Digitando o assunto no ofício
    sleep(1)

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "/html/body/p[3]/strong[2]")
    # texto_atual = celula.text

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, texto_substituir)

    # substituindo o nome na frase "Encaminhamos..."
    texto_substituir = f'Nº {numero_processo}'

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, texto_substituir)

    sleep(2)


    # Encontrar o elemento que contém o texto a ser modificado
    elemento = driver.find_element(By.XPATH,"//p[contains(text(),'Encaminhamos abaixo a solicitação de Data Especial de Colação de Grau de')]")

    # Substituir as variáveis no texto mantendo a formatação HTML
    novo_texto = f"Encaminhamos abaixo a solicitação de Data Especial de Colação de Grau de<strong> {aluno.nome[linha]} - CPF: {aluno.cpf[linha]} - MATRÍCULA: {aluno.matricula[linha]} </strong>recebida pela Central de Atendimento de forma remota em <strong>{aluno.data_solicitacao[linha]}</strong>."

    # Atualizar o texto no elemento
    driver.execute_script("arguments[0].innerHTML = arguments[1];", elemento, novo_texto)


    # DIGITANDO NA TABELA

    sleep(2)
    email = aluno.email[linha]
    # substituindo e-mail
    xpath_celula = "//html/body/table/tbody//tr[1]//td[2]"
    texto_substituir_email = aluno.email[linha]

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, xpath_celula)
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_email)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)
    ####################################################
    sleep(1)

    ####################################################
    # substituindo o nome
    texto_substituir_nome = aluno.nome[linha]

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[2]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_nome)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)
    ####################################################
    sleep(1)
    ####################################################

    # substituindo o cpf
    texto_substituir_cpf = cpf_tratado

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[3]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_cpf)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)
    ####################################################
    sleep(1)
    ####################################################
    # substituindo o RG
    xpath_celula = "//html/body/table/tbody//tr[4]//td[2]"
    texto_substituir_rg = aluno.rg[linha]

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, xpath_celula)
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_rg)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)
    ####################################################
    sleep(1)
    ####################################################
    ## substituindo a matricula

    texto_substituir_matricula = aluno.matricula[linha]
    #texto_substituir_matricula = aluno.contato_secundario[linha]

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[5]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_matricula)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)

    ####################################################
    sleep(1)
    ####################################################
    ## substituindo o telefone
    texto_substituir_telefone = aluno.telefone[linha]
    #texto_substituir_curso_anterior = aluno.curso_anterior[linha]

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[6]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_telefone)
    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)
    ####################################################
    sleep(1)
    ####################################################
    ## substituindo o logradouro
    #texto_substituir_matricula_anterior = aluno.matricula_anterior[linha]
    texto_substituir_logradouro = aluno.logradouro[linha]

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[7]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_logradouro)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)

    ####################################################
    sleep(1)
    ####################################################
    ## substituindo o numero/complemento
    # texto_substituir_curso_pretendido = aluno.curso_pretendido[linha]
    texto_substituir_numero_complemento = aluno.numero_complemento[linha]

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[8]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_numero_complemento)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)
    ####################################################
    sleep(1)
    ####################################################
    ## substituindo o bairro
    # texto_substituir_curso_pretendido = aluno.curso_pretendido[linha]
    texto_substituir_bairro = aluno.bairro[linha]

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[9]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_bairro)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)
    ####################################################
    sleep(1)
    ####################################################
    ## substituindo o cidade e estado
    # texto_substituir_curso_pretendido = aluno.curso_pretendido[linha]
    texto_substituir_estado_cidade = f'{aluno.cidade[linha]} - {aluno.estado[linha]}'

    driver.execute_script("window.scrollTo(0, 500);")

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[10]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_estado_cidade)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)
    ####################################################
    sleep(1)
    ####################################################
    ## substituindo o cep
    # texto_substituir_curso_pretendido = aluno.curso_pretendido[linha]
    texto_substituir_cep = aluno.cep[linha]

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[11]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_cep)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)
    ####################################################
    sleep(1)
    ####################################################
    ## substituindo o curso
    # texto_substituir_curso_pretendido = aluno.curso_pretendido[linha]
    texto_substituir_curso = aluno.curso[linha]

    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[12]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace("",texto_substituir_curso)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)
    ####################################################

    sleep(1)

    ####################################################        
    # digitando as considerações adicionais
    #texto_substituir_consideracoes = aluno.consideracoes[linha]
    texto_substituir_justificativa = aluno.justificativa[linha]
    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[13]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace('',texto_substituir_justificativa)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)

    sleep(1)
    driver.execute_script("window.scrollTo(0, 500);")



    texto_substituir_anexo = 'DOCUMENTO SEI Nº '
    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body/table/tbody//tr[14]//td[2]")
    texto_atual = celula.text

    # Realize a substituição
    novo_texto = texto_atual.replace('',texto_substituir_anexo)

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)

    celula.click()



    sleep(2)
    ## digitando o nome do operador
    #texto_substituir_operador = operador.nome_operador
    texto_substituir_operador = config('NOME_OPERADOR')
    # Localize a célula e seu texto atual
    celula = driver.find_element(By.XPATH, "//html/body//p[9]")
    texto_atual = celula.text

    # Realize a substituição
    sleep(1)
    novo_texto = texto_atual.replace("NOME",texto_substituir_operador.upper())

    # Use JavaScript para atualizar o conteúdo da célula
    driver.execute_script("arguments[0].innerText = arguments[1];", celula, novo_texto)


    # SALVANDO O DOCUMENTO
    sleep(3)
    actions = ActionChains(driver)

    # Pressione as teclas CTRL+ALT+S
    actions.key_down(Keys.CONTROL)
    actions.key_down(Keys.ALT)
    actions.send_keys("s")
    actions.key_up(Keys.CONTROL)
    actions.key_up(Keys.ALT)

    # Execute as ações
    actions.perform()
    sleep(3)


    ### incluindo o anexo

    # obtendo a data atual
    DATA_ATUAL =  datetime.date.today().strftime('%d/%m/%Y')

    driver.switch_to.window(driver.window_handles[0])

    # INCLUINDO ANEXO
    sleep(1)
    iframe = driver.find_element(By.XPATH,"//div[@id='divConteudo']/iframe[1]")
    driver.switch_to.frame(iframe)


    # clicando no numero do processo
    sleep(1)
    botao_numero_processo = driver.find_element(By.XPATH,"//div[@id='header']/div/a[@target='ifrVisualizacao'][1]")
    botao_numero_processo.click()

    driver.switch_to.window(driver.window_handles[0])
    sleep(1)
    iframe = driver.find_element(By.XPATH,"//div[@id='divInfraAreaTelaD']//div/iframe[2]")
    driver.switch_to.frame(iframe)

    # clicando em novo documento
    sleep(1)
    botao_incluir_doc = driver.find_element(By.XPATH,"//div//a/img[@title='Incluir Documento']")
    botao_incluir_doc.click()

    # Escolhendo o tipo de documento
    sleep(1)
    campo_tipo_doc = driver.find_element(By.XPATH,"//div/input[@id='txtFiltro']")
    campo_tipo_doc.send_keys('Externo')
    sleep(1)
    opcao_externo = driver.find_element(By.XPATH,"//table/tbody/tr/td/a[@class='ancoraOpcao']")
    opcao_externo.click()

    # registando o tipo de documento
    sleep(1)
    campo_tipo_doc = driver.find_element(By.XPATH,"//select[@id='selSerie']")
    campo_tipo_doc.click()
    campo_tipo_doc.send_keys('Anexo')
    sleep(1)
    campo_tipo_doc.send_keys(Keys.ENTER)

    # registando a data do documento
    sleep(1)
    campo_tipo_doc = driver.find_element(By.XPATH,"//div[@id='divSerieDataElaboracao']/input[@id='txtDataElaboracao']")
    campo_tipo_doc.click()
    sleep(1)
    campo_tipo_doc.send_keys(DATA_ATUAL)

    # clicando em formato
    sleep(1)
    botao_formato = driver.find_element(By.XPATH,"//fieldset[@id='fldFormato']//div/input[1]")
    botao_formato.click()

    # escolhendo o nível de acesso do documento
    sleep(1)
    radio_restrito = driver.find_element(By.ID,"optRestrito")
    radio_restrito.click()

    # definindo a hipótese legal
    sleep(1)
    botao_nivel_acesso_sei = driver.find_element(By.XPATH,"//div[@id='divNivelAcesso']/fieldset/select[@id='selHipoteseLegal']") 
    botao_nivel_acesso_sei.click()
    sleep(1)
    botao_nivel_acesso_sei.send_keys('Informação')
    sleep(1)
    botao_nivel_acesso_sei.send_keys(Keys.ENTER)
    sleep(1)

    # Clicando no botão escolhendo arquivo
    sleep(1)
    botao_escolher_arquivo = driver.find_element(By.XPATH,"//div[@id='divArquivo']/input[@id='filArquivo']")

    ##############################################



    ##### PEGANDO OS DOCUMENTOS A SEREM ANEXADOS #####

    pasta = r"D:\CentralDeAtendimento\Projetos\ProjetosCAT\ProjetoAtecipacaoColacaoDeGrauCAT\documentosDataEspecial"
    #pasta = r"C:\projetos\ProjetoAntecipacaoColacaoDeGrauCAT\documentosDataEspecial"


    # Loop para iterar sobre os arquivos
    arquivo = f'docs{linha}.pdf'


    # Constrói o caminho completo do arquivo
    caminho_completo = os.path.join(pasta, arquivo)

    botao_escolher_arquivo.send_keys(caminho_completo) 


    ### TESTE PARA OBTENÇÃO DE DOCUMENTOS ###

    # arquivo = f'docs.pdf'

    # caminho_completo = os.path.join(pasta, arquivo)

    # botao_escolher_arquivo.send_keys(caminho_completo) 

    ##### --------------------------------------- #####


    ###############################

    # clicando  em confirmar
    sleep(1)
    botao_confirmar =  driver.find_element(By.XPATH,"//div//div[@id='divInfraBarraComandosInferior']/button[@id='btnSalvar']")
    botao_confirmar.click()
    sleep(1)

    driver.switch_to.window(driver.window_handles[0])
    sleep(1)
    iframe = driver.find_element(By.XPATH,"//div[@id='divConteudo']/iframe[1]")
    driver.switch_to.frame(iframe)

    # pegando número sei do anexo
    # clicando no numero do anexo
    sleep(1)
    botao_numero_anexo = driver.find_element(By.XPATH,"//div[@id='divArvore']/div//a[4]/img")
    botao_numero_anexo.click()
    sleep(1)

    # pegando o numero do anexo
    numero_anexo = pyperclip.paste()

    # clicando no nome do oficio
    sleep(1)
    botao_confirmar =  driver.find_element(By.XPATH,"//div[@id='divArvore']/div//a[1]/img")
    botao_confirmar.click()
    sleep(1)

    # pegando o numero do anexo
    numero_oficio = pyperclip.paste()
    sleep(3)


    driver.switch_to.window(driver.window_handles[1])
    sleep(2)
    # Maximizando janela
    driver.maximize_window()

    #clicando no simbolo link sei
    sleep(2)
    botao_editar_doc =  driver.find_element(By.XPATH,"//body/form/div//div[3]/div/div/span[2]/span[6]/span[2]/a[3]/span[1]")
    botao_editar_doc.click()

    sleep(3)
    # digitando o numero do anexo
    botao_editar_doc =  driver.find_element(By.XPATH,"//body//table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/div/div/input")
    botao_editar_doc.send_keys(numero_anexo)
    sleep(1)

    # clicando em ok
    botao_ok = driver.find_element(By.XPATH,"//body//table/tbody/tr/td/div/table/tbody/tr[2]/td/table/tbody/tr/td/a")
    botao_ok.click()


    # SALVANDO O DOCUMENTO
    sleep(2)
    actions = ActionChains(driver)

    # Pressione as teclas CTRL+ALT+S
    actions.key_down(Keys.CONTROL)
    actions.key_down(Keys.ALT)
    actions.send_keys("s")
    actions.key_up(Keys.CONTROL)
    actions.key_up(Keys.ALT)

    # Execute as ações
    actions.perform()


    # ASSINANDO O DOCUMENTO
    sleep(2)
    actions2 = ActionChains(driver)

    # Pressione as teclas CTRL+ALT+S
    actions2.key_down(Keys.CONTROL)
    actions2.key_down(Keys.ALT)
    actions2.send_keys("a")
    actions2.key_up(Keys.CONTROL)
    actions2.key_up(Keys.ALT)

    # Execute as ações
    actions2.perform()


    # assinando documento
    driver.switch_to.window(driver.window_handles[2])
    sleep(2)
    campo_assinatura = driver.find_element(By.XPATH,"//body/div/div/div/form//div[2]//div[5]/input[@type='password']")
    # campo_assinatura.send_keys(operador.senha_operador)
    campo_assinatura.send_keys(config('SENHA_OPERADOR'))

    sleep(2)
    botao_assinatura = driver.find_element(By.XPATH,"//body/div/div/div/form//div/button[@name='btnAssinar']")
    botao_assinatura.click()


    # ### envio do e-mail de dentro do processo

    driver.switch_to.window(driver.window_handles[0])

    sleep(1)
    iframe = driver.find_element(By.XPATH,"//div[@id='divConteudo']/iframe[1]")
    driver.switch_to.frame(iframe)


    # clicando no numero do processo
    sleep(1)
    botao_numero_processo = driver.find_element(By.XPATH,"//div[@id='header']/div/a[@target='ifrVisualizacao'][1]")
    botao_numero_processo.click()


    sleep(5)
    # Clicando no ícone para abrir a janela envio de correspondência
    driver.switch_to.window(driver.window_handles[0])
    sleep(1)
    iframe = driver.find_element(By.XPATH,"//div[@id='divInfraAreaTelaD']//div/iframe[2]")
    driver.switch_to.frame(iframe)

    # clicando em novo documento
    sleep(1)
    botao_envio_email = driver.find_element(By.XPATH,"//*[@id='divArvoreAcoes']/a[10]")
    botao_envio_email.click()

    # Digitando o destinatário do e-mail
    driver.switch_to.window(driver.window_handles[1])
    sleep(1)


    campo_endereco_email = driver.find_element(By.XPATH,"//div[@id='divPara']/p/div[@id='s2id_hdnDestinatario']/ul/li/input")
    #campo_endereco_email.send_keys('gra.cdara@ufjf.br')
    campo_endereco_email.send_keys('carlosrodriguesjf@gmail.com')

    sleep(1)

    campo_endereco_email.send_keys(Keys.ENTER)

    # Digitando assunto do e-mail
    sleep(1)
    campo_assunto = driver.find_element(By.XPATH,"//div[@id='divAssuntoMensagem']/input[1]")
    campo_assunto.send_keys(f'Data Especial de Colação de Grau - {aluno.nome[linha]}')


    # Digitando o corpo do e-mail
    sleep(2)

    corpo_email = f''' À Pró-reitoria de Graduação da UFJF,

    Assunto: Data especial de Colação de Grau - {aluno.nome[linha]}
    Número do Processo SEI: {numero_processo}



        Prezado Coordenador(a) da CDARA, encaminhamos para sua ciência.



    {config('NOME_OPERADOR').upper()}

    Central de Atendimento da UFJF'''

    campo_corpo_email = driver.find_element(By.XPATH,"//div[@id='divAssuntoMensagem']/textarea")
    campo_corpo_email.send_keys(corpo_email)
    sleep(2)


    # Seleccionando os anexos
    driver.execute_script("window.scrollTo(0, 1500);")
    sleep(2)
    checkbox_anexos = driver.find_element(By.XPATH,"//div[@id='divInfraAreaTabela']/table//tbody/tr/th/a")
    checkbox_anexos.click()


    # Enviando o e-mail
    sleep(1)
    # botao_enviar_email = driver.find_element(By.XPATH,"//div[@id='divInfraBarraComandosInferior']//button[@name='btnEnviar']")
    # botao_enviar_email.click()                                         
    botao_cancelar_email = driver.find_element(By.XPATH,"//div[@id='divInfraBarraComandosInferior']//button[@name='btnCancelar']")
    botao_cancelar_email.click()   
    sleep(2)



    # # clicando no alerta de confirmação
    # alerta = driver.switch_to.alert
    # alerta.accept()
    # sleep(2)



    driver.switch_to.window(driver.window_handles[0])

    # clicando no botão enviar processo
    sleep(1)
    iframe = driver.find_element(By.XPATH,"//iframe[2][@id='ifrVisualizacao']")
    driver.switch_to.frame(iframe)

    sleep(1)
    botao_incluir_doc = driver.find_element(By.XPATH,"//div//a/img[@alt='Enviar Processo']")
    botao_incluir_doc.click()

    # pegando o numero de processo
    sleep(1)
    numero_processo = driver.find_element(By.ID,'selProcedimentos')
    opcao_selecionada = numero_processo.find_element(By.TAG_NAME,'option').text  
    numero_processo = opcao_selecionada.split(' - ')[0]

    # digitando no campo PROTOCOLO
    sleep(2)
    campo_unidades = driver.find_element(By.XPATH,"//div//input[@id='txtUnidade']")
    campo_unidades.send_keys("SEC-PROGRAD")
    sleep(1)

    campo_unidades.send_keys(Keys.DOWN)
    sleep(1)
    campo_unidades.send_keys(Keys.ENTER)

    sleep(1)
    checkbox_manter_aberto = driver.find_element(By.ID,"chkSinManterAberto")
    checkbox_manter_aberto.click()
    sleep(1)
    checkbox_manter_aberto = driver.find_element(By.ID,"chkSinEnviarEmailNotificacao")
    checkbox_manter_aberto.click()

    # #Enviando o processo
    # botao_enviar_processo = driver.find_element(By.XPATH,"//form//div[1]/button[@type='submit']")
    # botao_enviar_processo.click()


    sleep(5)
    driver.close()

    return email, numero_processo, numero_oficio


