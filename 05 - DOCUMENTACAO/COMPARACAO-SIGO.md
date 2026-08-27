# Comparação de Telefonia com SIGO

## Objetivo

Atualizar a aba `Planos` da planilha de telefonia usando a base mais recente do SIGO, de forma simples e sem criar abas ou aplicar cores.

## Antes de comparar

A automação remove filtros e ordenações aplicados na aba `Planos` e reexibe todas as linhas. Assim, toda a base é considerada.

## Como a pessoa é localizada

Para cada linha da planilha, a automação tenta localizar o cadastro no SIGO nesta ordem:

1. Nome normalizado.
2. Chapa, se não encontrar pelo nome.

Nos nomes, são removidos acentos, espaços duplicados e partículas como `de`, `da`, `do`, `das`, `dos` e `e`. A mesma normalização é usada na planilha de telefonia e no SIGO.

## Regras aplicadas

| Situação | Resultado |
| --- | --- |
| Encontrado no SIGO e ativo | Atualiza somente o CPF. Mantém o Status atual. |
| Encontrado no SIGO e desligado | Altera somente o Status para `Desligado`. |
| Não encontrado no SIGO por nome nem chapa | Altera o Status para `VERIFICAR`. |
| Chapa/CPF contém `FROTA` | Ignora a linha e preserva os dados atuais. |
| Chapa/CPF contém `FAMILIA` | Ignora a linha e preserva os dados atuais. |
| Chapa/CPF contém `FORA SIGO` | Não consulta o SIGO e define o Status como `ATIVA`. Use para terceiros ou linhas de setores sem cadastro no SIGO. |
| Status já é `ESTOQUE` | Ignora a linha e preserva os dados atuais. |

## O que a comparação não altera

A comparação não cria abas, não usa cores e não atualiza nome, chapa, linha, centro de custo ou data. Apenas CPF e Status podem ser modificados, conforme as regras acima.

## Resultado e segurança

Após cada execução, o menu mostra o total de linhas preenchidas verificadas, incluindo FROTA. Antes de salvar, é criado um backup em `04 - SAIDAS\BACKUPS`.
