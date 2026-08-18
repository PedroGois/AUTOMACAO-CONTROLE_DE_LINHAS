# Controle de Telefonia

Esta automação mantém a base de telefonia atualizada, consulta a situação dos colaboradores no SIGO e atualiza o Dashboard.

## O que ela faz

1. Copia as planilhas de origem para `01 - DADOS`.
2. Baixa a base atualizada do SIGO.
3. Compara as linhas de telefonia com os dados do SIGO.
4. Atualiza os arquivos de dados usados pelo Dashboard.

O processo automático não gera e-mails, cobranças ou arquivos separados por Centro de Custo.

## Antes de usar

- Instale as dependências uma única vez:

  ```powershell
  pip install -r requirements.txt
  ```

- Preencha o arquivo `.env` na raiz do projeto com as credenciais do SIGO:

  ```env
  SIGO_DOCUMENT=seu_documento
  SIGO_PASSWORD=sua_senha
  ```

As credenciais são necessárias para a atualização automática funcionar sem pedir dados na tela.

## Atualizar as planilhas de origem

Use `03 - BAT\00 - COPIAR DADOS PARA AUTOMACAO.bat` para copiar:

- `TELEFONIA.xlsx`
- `CONTATO CDC.xlsx`

para a pasta `01 - DADOS` do projeto.

Para agendar essa cópia a cada hora, execute uma vez:

```text
03 - BAT\AGENDAR COPIA DADOS A CADA HORA.bat
```

O log dessa cópia fica em `04 - SAIDAS\LOGS\copia_dados.log`.

## Atualizar o Dashboard

Para executar manualmente todo o fluxo de atualização, use:

```text
03 - BAT\06 - ATUALIZAR DASHBOARD AUTOMATICO.bat
```

Esse arquivo atualiza a base SIGO, compara a telefonia e gera os dados do Dashboard.

Para agendar esse fluxo a cada hora, execute uma vez:

```text
03 - BAT\AGENDAR ATUALIZACAO DASHBOARD A CADA HORA.bat
```

O log fica em `04 - SAIDAS\LOGS\atualizacao_dashboard.log`.

## Abrir o Dashboard

Abra `00 - DASHBOARD\index.html` no navegador.

Se o aviso “Base não gerada” aparecer, execute `06 - ATUALIZAR DASHBOARD AUTOMATICO.bat` e confira o log se houver erro.

## Regras da comparação

- Colaborador ativo no SIGO: linha fica como `ATIVA`.
- Colaborador desligado: linha fica como `DESLIGADO`.
- Não encontrado no SIGO: linha fica como `VERIFICAR`.
- Linhas de frota e estoque seguem suas regras próprias.
