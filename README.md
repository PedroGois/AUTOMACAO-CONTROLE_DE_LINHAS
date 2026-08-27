# Automação de Telefonia

## ⚠️ O problema

As linhas corporativas precisam ser conferidas com frequência. A base de telefonia pode ficar desatualizada, manter linhas vinculadas a pessoas desligadas ou ter pendências sem responsável definido.

Fazer essa conferência manualmente consome tempo e aumenta o risco de usar uma base antiga.

## ✅ A solução

Esta automação atualiza a base de telefonia e o dashboard em um único fluxo:

~~~text
Planilhas corporativas
        ↓
Consulta ao SIGO
        ↓
Comparação e atualização da telefonia
        ↓
Dashboard de Telefonia
~~~

O processo cria backups, registra logs e separa as pendências para facilitar a conferência.

## ▶️ Como utilizar

### Uso diário

1. Feche a planilha **TELEFONIA.xlsx**, caso esteja aberta.
2. Abra **MENU DA AUTOMACAO.bat**.
3. Escolha **1 — Atualizar dados e Dashboard Gerencial**.
4. Ao terminar, abra **00 - DASHBOARD\index.html** para conferir o resultado.

Essa é a opção de uso normal. Ela copia as planilhas corporativas, consulta o SIGO, compara os dados e atualiza o dashboard.

### Outras opções do menu

- **2 — Exportar base atualizada para a planilha original**  
  Devolve a planilha atualizada para a pasta corporativa. Pede confirmação e cria backup da versão anterior.

- **3 — Separar linhas VERIFICAR por CDC**  
  Gera arquivos de conferência organizados por centro de custo.

- **4 — Gerar e-mails EML da VIVO e TIM**  
  Cria rascunhos de cobrança das linhas pendentes.

- **5 — Executar fluxo completo**  
  Atualiza dados, separa pendências e gera e-mails em sequência.

- **6 — Anonimizar base do Dashboard para Git**  
  Prepare a base do dashboard para envio ao Git depois de concluir a conferência. A opção anonimiza nomes, linhas, CPF, chapas, CDCs e identificadores de aparelhos.

### Atualização automática

Execute uma única vez:

~~~text
03 - BAT\AGENDAR\AGENDAR ATUALIZACAO COMPLETA A CADA HORA.bat
~~~

Apesar do nome do arquivo, ele está configurado para executar o fluxo completo a cada **30 minutos**.

No Agendador de Tarefas, o estado **Pronto** significa que a tarefa está configurada e aguardando o próximo horário. **Em execução** aparece somente durante o processamento.

### Primeira configuração

1. Instale Python 3.11 ou mais recente, com a opção **Add Python to PATH**.
2. Instale as dependências:

   ~~~powershell
   python -m pip install -r requirements.txt
   ~~~

3. Crie o arquivo **.env** na raiz com suas credenciais do SIGO:

   ~~~env
   SIGO_DOCUMENT=seu_documento
   SIGO_PASSWORD=sua_senha
   ~~~

4. Confirme que existem os arquivos corporativos **telefonia.xlsx** e **CONTATO CDC.xlsx** em **Ti\Telefonia e Internet\TELEFONIA**.

## 📌 Regras

- A consulta ao SIGO é somente de leitura.
- A comparação atualiza CPF e Status conforme as regras de negócio.
- Linhas podem ficar como **ATIVA**, **ESTOQUE**, **DESLIGADO** ou **VERIFICAR**.
- Linhas marcadas como **FROTA**, **FAMILIA** e **FORA SIGO** seguem regras próprias. Consulte [as regras detalhadas da comparação](05%20-%20DOCUMENTACAO/COMPARACAO-SIGO.md).
- A aba **Aparelhos** do dashboard é apenas de consulta; ela não altera dados no SIGO.
- Nunca envie o arquivo **.env**, credenciais ou dados operacionais ao Git.

## ℹ️ Informações importantes

### Onde conferir o resultado

- Dashboard: **00 - DASHBOARD\index.html**
- Base SIGO atual: **04 - SAIDAS\BASE_SIGO.xlsx**
- Resultado da comparação: **04 - SAIDAS\resultado_comparacao.txt**
- Log do fluxo: **04 - SAIDAS\LOGS\atualizacao_dashboard.log**
- Backups: **04 - SAIDAS\BACKUPS**

### Estrutura do projeto

~~~text
AUTO - CONTROLE DE LINHAS/
├── MENU DA AUTOMACAO.bat         # ponto de entrada manual
├── 00 - DASHBOARD/               # painel e seu repositório próprio
├── 01 - DADOS/                   # cópias das planilhas corporativas
├── 02 - SCRIPTS/                 # integração, comparação e geração
├── 03 - BAT/                     # ações do menu e agendamento
├── 04 - SAIDAS/                  # base SIGO, resultados, logs e backups
└── 05 - DOCUMENTACAO/            # regras detalhadas da automação
~~~

### Documentação relacionada

- Regras completas de comparação: **05 - DOCUMENTACAO\COMPARACAO-SIGO.md**
- Documentação do dashboard: **00 - DASHBOARD\README.MD**
