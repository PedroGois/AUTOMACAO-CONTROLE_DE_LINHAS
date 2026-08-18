---
name: controle-de-linhas
description: Use esta skill ao planejar, criar, analisar ou melhorar o sistema de controle de linhas corporativas, custos, atualizações e colaboradores.
---

# Contexto do projeto: Controle de linhas

Este projeto é um controle de linhas corporativas. O objetivo é permitir que a empresa saiba:

- Onde está cada linha e qual é o seu status.
- Qual colaborador é responsável ou utiliza cada linha.
- Quanto é pago por cada linha e o custo total mensal.
- Quais linhas estão sem atualização há mais de 30 dias.
- Quais linhas estão vinculadas a colaboradores desligados no SIGO.
- Quanto a empresa paga, por mês, pelas linhas desatualizadas ou vinculadas a colaboradores desligados.

## Regras de negócio iniciais

- Uma linha deve ter, no mínimo: número, operadora, valor mensal, status, colaborador responsável, data da última atualização e observações.
- Uma linha é considerada desatualizada quando a última atualização ocorreu há mais de 30 dias.
- Uma linha deve ser sinalizada quando o colaborador associado estiver desligado no SIGO.
- O painel deve exibir:
  - total de linhas;
  - custo mensal total;
  - quantidade e custo das linhas desatualizadas;
  - quantidade e custo das linhas associadas a colaboradores desligados;
  - lista das linhas que exigem ação.
- Não assumir que uma linha vinculada a colaborador desligado deve ser cancelada: sinalizar para revisão humana.
- Não inventar dados, regras ou integrações com o SIGO. Quando uma informação estiver ausente, perguntar ou registrar como pendência.

## Diretrizes de trabalho

- Antes de implementar, apresentar um plano simples: telas, dados necessários e regras de cálculo.
- Priorizar uma primeira versão simples, clara e auditável.
- Exibir valores em reais (R$) e datas no formato brasileiro.
- Criar filtros por status, operadora, colaborador, atualização e situação no SIGO.
- Toda alteração de status, valor ou responsável deve guardar data e responsável pela alteração.
- Antes de mudanças grandes, pedir confirmação.
- Ao concluir uma tarefa, informar o que foi criado, quais dados ainda faltam e como validar o resultado.
## Integração com SIGO

- O projeto possui um script Python responsável por consultar o SIGO e obter a situação dos colaboradores.
- O agente deve reutilizar e preservar esse fluxo Python existente; não substituir a integração sem necessidade.
- A consulta deve retornar, no mínimo, um identificador confiável do colaborador e sua situação, incluindo se está desligado.
- O cruzamento deve usar um identificador estável, preferencialmente matrícula/ID do colaborador, e não apenas nome.
- Linhas vinculadas a colaboradores desligados no SIGO devem ser sinalizadas para revisão e entrar no cálculo de custo mensal potencialmente indevido.
- A integração com o SIGO deve ser somente de leitura, salvo autorização explícita.
- Nunca incluir senhas, tokens ou credenciais do SIGO no código, no repositório ou em mensagens. Usar variáveis de ambiente ou o mecanismo seguro que já existir no projeto.
- Antes de alterar o script Python de integração, analisar o funcionamento atual e apresentar um plano de mudança.

