
# 📄 Automação de Protocolos no Sistema Eletrônico de Informações (SEI)

## 👤 Autor e Contato
- **Nome**: Carlos Eduardo Rodrigues  
- **E-mail**: carlosrodriguesjf@gmail.com  
- **LinkedIn**: https://www.linkedin.com/in/carlosrodriguesjf
- **GitHub**: https://github.com/carlosrodriguesjf

---

## 📌 Descrição
Este projeto é um script em **Python** que automatiza o protocolo de processos no **Sistema Eletrônico de Informações (SEI)** da Universidade Federal de Juiz de Fora.  

Ele realiza automaticamente:
- Coleta de dados via **Google Sheets API**  
- Abertura e preenchimento de processos no SEI com **Selenium**  
- Anexação de documentos ao processo  
- Envio de e-mails automáticos via **PyAutoGUI**  
- Registro do número do processo e do documento em planilhas do Google  

O foco é reduzir o tempo de execução, minimizar erros manuais e padronizar o fluxo institucional.

---

## 🔍 Considerações Iniciais 
- Os dados usados neste repositório são **fictícios ou anonimizados**.  
- É necessário possuir credenciais de acesso ao Google API e ao SEI para executar.  
- O fluxo automatizado foi projetado para o processo de **Data Especial de Colação de Grau**.

---

## 🎯 Objetivos
- Automatizar a abertura e preenchimento de processos no SEI.  
- Anexar documentos automaticamente.  
- Registrar dados no Google Sheets.  
- Enviar confirmação por e-mail para o solicitante.  

---

## 📂 Estrutura do Projeto
```
📁 documentosDataEspecial
│── 
│── main.py                  # Módulo principal
│── leitura_googlesheets.py  # Leitura da planilha
│── processo_sei.py          # Automação do SEI
│── gravacao_googlesheets.py # Gravação na planilha
│── envio_email.py           # Envio de e-mails automáticos
│── README.md                # Documentação do projeto
│── requirements.txt         # Dependências do projeto
│── token.json               # Credenciais Google (não versionado)
│── credentials.json         # Credenciais Google (não versionado)
│── tratamento_cpf           # Tratamento de cpf
│── .env                     # Variáveis de ambiente (não versionado)
```

---

## 📊 Dados Utilizados
- **Google Sheets** com dados dos alunos:
  - Nome, CPF, RG, Matrícula, E-mail, Telefone, Endereço, Curso, Justificativa
- **Documentos PDF** para anexar ao processo
- **Variáveis de ambiente** no `.env`:
  ```
  CPF_OPERADOR=seu_cpf
  SENHA_OPERADOR=sua_senha
  NOME_OPERADOR=Seu Nome Completo
  ```

---

## 🛠️ Tecnologias e Ferramentas
- Python  
- Pandas  
- Selenium  
- PyAutoGUI  
- Pyperclip  
- Google Sheets API  
- Decouple  
- Chrome WebDriver  

---

## 🚀 Como Executar o Projeto
1. **Clone o repositório**
   ```bash
   git clone https://github.com/carlosrodriguesjf/ProjetoAntecipacaoColacaoDeGrauCAT
   cd automacao_protocolo_sei
   ```

2. **Crie o ambiente virtual e instale as dependências**
   ```bash
   python -m venv venv
   source venv/bin/activate  
   # Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. **Configure as credenciais**
   - Crie o arquivo `.env` com:
     ```
     CPF_OPERADOR=seu_cpf
     SENHA_OPERADOR=sua_senha
     NOME_OPERADOR=nome_completo
     ```
   - Adicione os arquivos `credentials.json` e `token.json` do Google API.

4. **Execute**
   ```bash
   python scripts/main.py
   ```

---

## 📈 Resultados e Benefícios
- Tempo de protocolo reduzido de horas para minutos.  
- Padronização de documentos e textos.  
- Registro automático no Google Sheets.  
- Notificação automática ao solicitante.  

---

## 🎥 Demonstração
*(Insira aqui um GIF ou link para vídeo do script rodando, ex: YouTube)*  


---

## 🔮 Melhorias Futuras
- Interface gráfica amigável.  
- Integração com outros sistemas da instituição.  
- Logs de execução mais detalhados.  
- Adaptação para diferentes tipos de processos no SEI.  


---

## 📄 Licença
Este projeto é para fins  institucionais.  
Não deve ser utilizado para burlar políticas de uso do SEI ou expor dados sensíveis.
