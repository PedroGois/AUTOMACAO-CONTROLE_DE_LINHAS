# 📱 Controle de Linhas - Automação de Gestão de Telefonia & Auditoria SIGO

> Sistema automatizado para conferência, auditoria de linhas telefônicas corporativas, cruzamento com a base de Recursos Humanos (SIGO), detecção de custos indevidos com colaboradores desligados, geração de cobranças por Centro de Custo e Dashboard Gerencial.

---

## 🎯 Objetivo do Projeto

Reduzir custos e otimizar a gestão de telecomunicações corporativas através da conferência automatizada entre o inventário de linhas ativas (VIVO, TIM, etc.) e a base de colaboradores do sistema de RH/SIGO.

### Principais Benefícios:
- **Identificação Imediata de Desperdício**: Localiza linhas ativas atreladas a colaboradores com status `DESLIGADO`.
- **Cruzamento Inteligente**: Sistema de matching com priorização estrita: `Chapa/Matrícula` ➔ `CPF` ➔ `Nome Normalizado` (com remoção de acentos e partículas).
- **Tratamento de Exceções**: Suporte a regras especiais para `FROTA`, `FAMILIA`, `FORA SIGO` e `ESTOQUE`.
- **Auditoria & Cobrança por Centro de Custo (CDC)**: Segmenta automaticamente pendências e gera rascunhos de e-mail (`.eml` ou Outlook) prontos para validação com gestores.
- **Dashboard Gerencial Interativo**: Painel visual em HTML com indicadores de custos, status e histórico de variações.

---

## 🗂️ Estrutura do Projeto

```text
TESTE TELEFONIA/
├── 01 - DADOS/                         # Bases de entrada (planilhas de telefonia e contatos)
│   ├── TELEFONIA.xlsx                  # Base principal de linhas (ignorado no Git)
│   └── CONTATO CDC-TESTE.xlsx          # Contatos dos responsáveis por cada CDC (ignorado no Git)
│
├── 02 - SCRIPTS/                       # Scripts em Python com a lógica de negócio
│   ├── baixar_base_sigo.py             # Integração via API para extração da base de RH
│   ├── comparar_telefonia_sigo.py      # Motor de cruzamento de dados e auditoria
│   ├── separar_verificar_por_centro_custo.py # Segmentação de pendências por CDC
│   ├── cobranca_verificar.py           # Geração de rascunhos de e-mail de cobrança
│   └── gerar_dashboard.py              # Compilação do dashboard gerencial em HTML
│
├── 03 - BAT/                           # Executáveis Windows Batch para cada etapa
│   ├── 01 - ATUALIZAR BASE SIGO.bat
│   ├── 02 - COMPARAR COM SIGO.bat
│   ├── 03 - SEPARAR POR CDC.bat
│   ├── 04 - GERAR E-MAILS EML.bat
│   └── 05 - GERAR DASHBOARD.bat
│
├── 04 - SAIDAS/                        # Relatórios, backups e arquivos gerados (ignorado no Git)
│   ├── BACKUPS/                        # Backups automáticos criados antes de cada alteração
│   ├── HISTORICO_SIGO/                 # Snapshots das bases SIGO baixadas
│   ├── VERIFICAR POR CENTRO DE CUSTO/  # Planilhas segmentadas por gestor/CDC
│   ├── COBRANCA E-MAILS/               # Rascunhos gerados (.eml)
│   └── resultado_comparacao.txt        # Métricas da última auditoria
│
├── 05 - DOCUMENTACAO/                  # Documentos técnicos e fluxogramas
│   ├── COMPARACAO-SIGO.md              # Regras detalhadas do cruzamento
│   └── FLUXO AUTOMAÇÃO LINHA.png       # Diagrama de fluxo do processo
│
├── MENU DA AUTOMACAO.bat               # Menu principal interativo (CLI)
├── DASH.html                           # Dashboard HTML gerado
├── .gitignore                          # Proteção contra vazamento de dados sensíveis
├── .env.example                        # Exemplo de configuração de variáveis de ambiente
└── README.md                           # Este documento
```

---

## ⚙️ Pré-requisitos e Instalação

1. **Python 3.8+** instalado no Windows ([python.org](https://www.python.org/)).
2. Instalar as bibliotecas necessárias:
   ```bash
   pip install pandas openpyxl
   ```

---

## 🚀 Como Utilizar

A forma mais simples e recomendada de operar o sistema é pelo **Menu Interativo**:

1. Feche as planilhas em `01 - DADOS` (o Excel não deve travar os arquivos).
2. Dê um duplo clique em `MENU DA AUTOMACAO.bat`.
3. Escolha a opção desejada:

```text
======================================================
         MENU - AUTOMACAO DE TELEFONIA
======================================================
 1 - Atualizar a base SIGO
 2 - Comparar Telefonia com SIGO
 3 - Separar linhas VERIFICAR por CDC
 4 - Gerar e-mails EML da VIVO
 5 - Gerar e-mails EML da TIM
 6 - Executar fluxo completo
 7 - Atualizar Dashboard Gerencial

 0 - Sair
```

### 🔁 Fluxo Completo (Opção 6):
Executa sequencialmente o download da base SIGO ➔ comparação de dados ➔ separação por CDC ➔ geração de e-mails para a operadora escolhida ➔ atualização do Dashboard final.

---

## 🧠 Regras de Negócio e Cruzamento

| Situação Encontrada | Ação Aplicada |
| :--- | :--- |
| **Ativo no SIGO** | Atualiza/valida CPF e mantém status ativo |
| **Desligado no SIGO** | Atualiza status para `DESLIGADO` e computa o custo desperdiçado |
| **Não Encontrado no SIGO** | Define status como `VERIFICAR` para auditoria manual com o CDC |
| **`FROTA` / `FAMILIA`** | Linha especial mantida intacta |
| **`FORA SIGO`** | Define status como `ATIVA` sem consultar a base de RH |
| **`ESTOQUE`** | Preserva a linha em estoque |

> 🛡️ **Segurança em primeiro lugar**: Antes de qualquer gravação na planilha principal, um backup com carimbo de data e hora é salvo automaticamente em `04 - SAIDAS/BACKUPS/`.

---

## 🔒 Segurança e LGPD (Proteção de Dados Sensíveis)

Este repositório foi configurado para **não expor nenhum dado confidencial ou pessoal** (nomes de funcionários, números de telefone, CPFs, matrículas ou custos):

- O arquivo `.gitignore` bloqueia o envio de todas as planilhas reais em `01 - DADOS/`, pastas de `04 - SAIDAS/`, backups e credenciais `.env`.
- **Nunca comite senhas ou CPFs no código.** O script de integração com o SIGO solicita credenciais de forma mascarada via terminal ou por variáveis de ambiente locais.

---

## 📊 Dashboard Gerencial (`DASH.html`)

O sistema conta com um compilador que gera uma interface visual moderna e responsiva (`DASH.html`), permitindo visualizar:
- Total de linhas ativas vs. inativas vs. pendentes.
- Custo mensal consolidado por operadora e por status.
- Histórico comparativo de evolução dos custos.
