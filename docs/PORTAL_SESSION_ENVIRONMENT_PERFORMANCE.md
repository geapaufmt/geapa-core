# Sessão do Portal: ambiente e desempenho

## Causa raiz

O Portal HOMOLOG publicado usava frontend HOMOLOG/DEV, porém o snapshot do
backend Apps Script declarava `ambienteDadosV2: PROD`. Ao mesmo tempo, partes
do Core ignoravam as opções recebidas e o leitor operacional de domínios usava
DEV como padrão fixo. O resultado era uma execução híbrida: configuração,
permissões, sessão e bases V2 podiam ser resolvidas em ambientes diferentes.

O mesmo caminho afetava login por código, Minha situação e telas
administrativas. Além disso:

- Minha situação resolvia a sessão no Portal e novamente dentro do Core;
- em caso de falha, o fallback repetia a consulta cara;
- o leitor operacional abria e lia todas as abas do domínio;
- caches de Portal e configuração do Core não incluíam o ambiente;
- exceções do Core eram convertidas em `null`.

## Correção

- o ambiente é obrigatório nos contratos Portal -> Core;
- HOMOLOG usa `DEV` e valida coerência entre os dois marcadores;
- caches possuem o ambiente na chave;
- configurações, perfis e permissões são lidos pela linha do Registry do
  ambiente solicitado, sem fallback cruzado;
- a sessão é memoizada por execução e reutilizada em Minha situação;
- os contratos pontuais leem somente as abas necessárias;
- planilhas e registros já lidos são reutilizados na mesma execução;
- falhas preservam `errorCode`, etapa e `traceId` seguros;
- leituras iguais no navegador compartilham a mesma Promise pendente;
- cada rota autenticada possui atualização isolada da aba.

## Contagem determinística com adaptadores

| Fluxo frio | Antes | Depois |
| --- | ---: | ---: |
| Resoluções de sessão em Minha situação | 2 | 1 |
| Leituras de abas estimadas em Minha situação | até 67 | 14 |
| Resoluções de sessão na listagem administrativa | 2 | 1 |
| Leituras de abas estimadas na listagem administrativa | até 67 | 14 |
| Segunda leitura da mesma aba na mesma execução | 1 | 0 |
| Abertura repetida da mesma planilha/ambiente | possível | 0 |

As contagens anteriores derivam do caminho publicado: Pessoas (16 abas),
Vigências (7 abas), três fontes de acesso e resoluções repetidas. As contagens
novas são verificadas por adaptadores que contabilizam abertura e leitura.

Não há medição real pós-correção porque esta tarefa não autoriza publicação em
HOMOLOG. Após publicar, os metadados `meta.trace` permitem medir cada etapa e o
tempo total sem registrar e-mail, RGA, CPF ou token.

## Publicação futura em HOMOLOG

1. Mesclar/publicar primeiro a nova versão imutável do Core.
2. Fixar essa versão no branch de release HOMOLOG do Portal.
3. Gerar `config.js` com `npm.cmd run config:homolog`.
4. Executar os testes listados no PR do Portal.
5. Executar `clasp push --force` somente no projeto compartilhado e criar uma
   versão Apps Script nova para atualizar apenas o deployment HOMOLOG.
6. Publicar somente o canal Firebase Hosting HOMOLOG.
7. Validar login por código, Minha situação e telas administrativas, anotando
   `traceId`, `tempoTotalMs` e etapas.

O deployment PROD anterior deve permanecer inalterado. O rollback consiste em
reapontar apenas HOMOLOG para a versão Apps Script e o release Hosting
anteriores.
