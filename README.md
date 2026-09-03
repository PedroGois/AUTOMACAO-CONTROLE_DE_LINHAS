# Controle de Linhas

## Objetivo

Esta automação confere as linhas corporativas com o SIGO, destaca pendências e atualiza um dashboard gerencial. Ela reduz a conferência manual e facilita a identificação de linhas sem responsável, em estoque ou vinculadas a pessoas desligadas — casos que podem gerar custo desnecessário.

## Como funciona

1. Copia as planilhas corporativas para uma área local protegida.
2. Consulta a base do SIGO e compara os cadastros.
3. Atualiza nome, CPF e status quando as regras permitem.
4. Gera o dashboard, relatórios por centro de custo e rascunhos de cobrança quando solicitados.

O painel é apenas para consulta; alterações operacionais continuam na planilha corporativa. Antes de cancelar ou transferir uma linha, valide o caso com a área responsável.

## Uso diário

1. Feche a planilha de telefonia.
2. Abra `executar.bat` e escolha **1 — Atualizar dados e dashboard**.
3. Ao final, abra `dashboard\\index.html`.

O menu também permite exportar a base atualizada, separar pendências por centro de custo, gerar rascunhos EML, anonimizar o painel e configurar a atualização recorrente.

## Primeira configuração

1. Instale Python 3.11 ou superior.
2. Instale as dependências:

   ```powershell
   python -m pip install -r requirements.txt
   ```

3. Copie `config\\.env.example` para `config\\.env` e informe as credenciais do SIGO.
4. Confirme o acesso às planilhas corporativas em `Telefonia e Internet\\TELEFONIA`.

## Estrutura

```text
├── src/                 automação em Python
├── scripts/             cópia, exportação e agendamento
├── dashboard/           painel web (repositório próprio incorporado)
├── config/              modelo e configuração local
├── data/
│   ├── entrada/         cópias locais das planilhas corporativas
│   └── saidas/          relatórios, logs e backups locais
├── executar.bat         ponto de entrada
├── requirements.txt     dependências Python
└── README.md            esta documentação
```

## Arquivos principais

- `executar.bat`: uso diário e tarefas auxiliares.
- `src/comparar_telefonia_sigo.py`: regras de comparação com o SIGO.
- `src/gerar_dashboard.py`: base de dados do painel.
- `scripts/agendar_atualizacao.ps1`: agenda a atualização a cada 30 minutos e após o logon.

## Cuidados

- `config/.env`, planilhas, logs, backups e resultados são locais e ignorados pelo Git.
- Nunca publique dados reais do dashboard. Use a opção de anonimização e revise o resultado antes de enviar ao Git.
- A consulta ao SIGO é somente de leitura; a comparação altera a cópia local de `TELEFONIA.xlsx` e cria backup.

## Detalhes técnicos

As regras de comparação priorizam chapa, CPF e nome normalizado. Linhas de frota, família, fora do SIGO e estoque seguem exceções preservadas no código. As dependências estão em `requirements.txt`.
