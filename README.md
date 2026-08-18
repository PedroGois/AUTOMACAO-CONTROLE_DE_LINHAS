# 📱 Controle de Telefonia & Auditoria SIGO

> Automação para auditoria de linhas corporativas (VIVO/TIM), detecção de cobranças indevidas de colaboradores desligados, envio de cobranças por Centro de Custo e Dashboard de custos.

---

## 🎯 O que o sistema faz?

1. **Cruza dados com o RH (SIGO)**: Identifica se quem está usando a linha ainda trabalha na empresa.
2. **Corta desperdícios**: Detecta linhas ativas de funcionários **Desligados**.
3. **Cobra os gestores (CDC)**: Separa as pendências e cria rascunhos de e-mail (`.eml`) prontos para envio.
4. **Dashboard de Custos**: Painel visual no navegador para acompanhar valores, histórico e variações.

---

## ⚠️ IMPORTANTE: Setup Inicial

### 1. Instalar dependências Python
Ao baixar o projeto pela primeira vez, execute:
```powershell
pip install -r requirements.txt
```

### 2. Configurar dados reais
Este repositório inclui **templates de teste** para você entender o fluxo:
- `01 - DADOS/TELEFONIA-TESTE.xlsx` → Use como referência
- `01 - DADOS/CONTATO CDC-TESTE.xlsx` → E-mails dos gestores

**Para usar seus dados reais:**
1. Obtenha a planilha de telefonia (VIVO/TIM)
2. Renomeie para `TELEFONIA.xlsx` ou `CONTATO CDC.xlsx`
3. Coloque na pasta `01 - DADOS/`

> ⚠️ **Os scripts procuram primeiro pelos templates (`-TESTE`). Para usar dados reais, é essencial renomear ou modificar os nomes nos scripts.**

### 3. Configurar credenciais SIGO
Crie um arquivo `.env` na raiz do projeto:
```env
SIGO_DOCUMENT=seu_cpf_aqui
SIGO_PASSWORD=sua_senha_aqui
```
> ⚠️ **Nunca comite o `.env` preenchido! Ele está no `.gitignore` para proteger suas credenciais.**

---

## 🗂️ Estrutura do Projeto

```text
TESTE TELEFONIA/
│
├── 01 - DADOS/                         # Planilhas de entrada
│   ├── TELEFONIA-TESTE.xlsx            # ⭐ Template: Base geral de linhas e valores
│   ├── CONTATO CDC-TESTE.xlsx          # ⭐ Template: E-mails dos gestores por CDC
│   ├── TELEFONIA.xlsx                  # Seus dados reais (não será commitado)
│   └── CONTATO CDC.xlsx                # Seus dados reais (não será commitado)
│
├── 02 - SCRIPTS/                       # Códigos em Python (Automação)
│   ├── 1_baixar_base_sigo.py           # 1. Baixa a base atualizada do RH
│   ├── 2_comparar_telefonia_sigo.py    # 2. Cruza Telefonia x SIGO
│   ├── 3_separar_verificar_por_cdc.py  # 3. Divide pendências por gestor
│   ├── 4_cobranca_verificar.py         # 4. Cria os e-mails de cobrança
│   ├── 5_gerar_dashboard.py            # 5. Atualiza os dados do painel
│   └── requirements.txt                # Dependências Python
│
├── 03 - BAT/                           # Atalhos rápidos (clique duplo)
│   ├── 01 a 05 - Scripts individuais
│   └── MENU DA AUTOMACAO.bat           # Menu principal unificado
│
├── 04 - SAIDAS/                        # Arquivos gerados automaticamente (não commitados)
│   ├── BACKUPS/                        # Cópias de segurança automáticas
│   ├── VERIFICAR POR CDC/              # Planilhas divididas por gestor
│   ├── COBRANCA E-MAILS/               # E-mails prontos para envio (.eml)
│   └── HISTORICO_SIGO/                 # Histórico de consultas
│
├── 00 - DASHBOARD/                     # Painel visual
│   ├── DASH.html                       # Abra no navegador para ver gráficos
│   ├── assets/                         # CSS e JavaScript
│   │   ├── style.css                   # Cores e design do painel
│   │   └── app.js                      # Filtros, gráficos e cálculos
│   └── dados/                          # Dados que abastecem o painel (gerado)
│
├── .env                                # Credenciais SIGO (não será commitado)
├── .gitignore                          # Configuração de arquivos ignorados
├── requirements.txt                    # Dependências Python
└── README.md                           # Guia de uso
```

**⭐ Nota sobre Templates:**
- Os scripts procuram primeiro pelos arquivos `-TESTE` (templates)
- Para usar seus dados reais, renomeie para `TELEFONIA.xlsx` e `CONTATO CDC.xlsx`
- Arquivos reais não são commitados (segurança)

---

## 🚀 Como Usar

O jeito mais fácil é pelo **Menu Principal**:

1. Feche o Excel caso esteja com as planilhas abertas.
2. Dê duplo clique em `MENU DA AUTOMACAO.bat`.
3. Escolha o que deseja fazer:

* **Opção 6 (Fluxo Completo)**: Roda tudo do início ao fim (baixa RH ➔ audita ➔ divide por CDC ➔ gera e-mails ➔ atualiza painel).
* **Opção 7 (Atualizar Dashboard)**: Apenas processa a planilha e atualiza o `DASH.html`.

Para ver os gráficos e custos, basta dar duplo clique no arquivo **`DASH.html`**.

---

## 🧠 Como funciona a auditoria?

| O que encontrou no RH? | O que o sistema faz? |
| :--- | :--- |
| **Ativo no SIGO** | Mantém a linha como `ATIVA`. |
| **Desligado no SIGO** | Marca como `DESLIGADO` (alerta de corte/economia). |
| **Não encontrado** | Marca como `VERIFICAR` (vai para cobrança com o gestor). |
| **FROTA / ESTOQUE** | Regras especiais mantidas sem alteração. |

> 🛡️ **Backup Automático**: O sistema nunca altera sua planilha sem antes salvar uma cópia com data e hora na pasta `04 - SAIDAS/BACKUPS/`.
