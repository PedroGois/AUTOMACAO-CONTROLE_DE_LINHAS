# 📱 Controle de Telefonia & Auditoria SIGO

> Automação para auditoria de linhas corporativas (VIVO/TIM), detecção de cobranças indevidas de colaboradores desligados, envio de cobranças por Centro de Custo e Dashboard de custos.

---

## 🎯 O que o sistema faz?

1. **Cruza dados com o RH (SIGO)**: Identifica se quem está usando a linha ainda trabalha na empresa.
2. **Corta desperdícios**: Detecta linhas ativas de funcionários **Desligados**.
3. **Cobra os gestores (CDC)**: Separa as pendências e cria rascunhos de e-mail (`.eml`) prontos para envio.
4. **Dashboard de Custos**: Painel visual no navegador para acompanhar valores, histórico e variações.

---

## 🗂️ Estrutura do Projeto

```text
TESTE TELEFONIA/
│
├── 01 - DADOS/                         # Planilhas de entrada
│   ├── TELEFONIA.xlsx                  # Base geral de linhas e valores
│   └── CONTATO CDC-TESTE.xlsx          # E-mails dos gestores por CDC
│
├── 02 - SCRIPTS/                       # Códigos em Python (Automação)
│   ├── baixar_base_sigo.py             # 1. Baixa a base atualizada do RH
│   ├── comparar_telefonia_sigo.py      # 2. Cruza Telefonia x SIGO
│   ├── separar_verificar_por_cdc.py    # 3. Divide pendências por gestor
│   ├── cobranca_verificar.py           # 4. Cria os e-mails de cobrança
│   └── gerar_dashboard.py              # 5. Atualiza os dados do painel
│
├── 03 - BAT/                           # Atalhos rápidos (clique duplo)
│   ├── 01 a 05 - Scripts individuais
│   └── MENU DA AUTOMACAO.bat           # Menu principal unificado
│
├── 04 - SAIDAS/                        # Arquivos gerados automaticamente
│   ├── BACKUPS/                        # Cópias de segurança automáticas
│   ├── VERIFICAR POR CDC/              # Planilhas divididas por gestor
│   ├── COBRANCA E-MAILS/               # E-mails prontos para envio (.eml)
│   └── dados_dashboard.js / .json      # Dados que abastecem o painel
│
├── assets/                             # Visual do Dashboard
│   ├── style.css                       # Cores e design do painel
│   └── app.js                          # Filtros, gráficos e cálculos
│
├── DASH.html                           # Painel Gerencial (abra no navegador)
└── README.md                           # Guia de uso
```

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
