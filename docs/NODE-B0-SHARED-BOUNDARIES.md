# Fronteiras B0 e possível extração em B1

Análise, sem refactor de domínio e sem implantação.

## Duplicação

O B0 original copiava 194 fontes v43 para `base/core/`. A lista exata está em
`docs/baselines/core-dev-v43.json`. Nesta preservação, `materialize-baseline.cjs`
lê essas fontes do commit v43 e verifica todos os hashes. O cache local é
ignorado: **zero cópia Core permanente adicional**. O WIP de HEAD não é carregado.

Há código de teste em `prototype/node-command-executor/test-support/`: runtime em
memória, geradores Members/ACT, cenários e compatibilidade de queries. Ele era
uma dependência entre workspaces; agora está versionado com caminhos relativos.
Não é implementação institucional concorrente, cadastro ou snapshot importado.

## Separação candidata

| Lógica existente | Destino conceitual em B1 | Cuidado |
|---|---|---|
| `shared/core/10_shared_core`, schemas, lifecycle, transaction plans | Shared contracts/domain | Manter códigos de erro, digest integral, estado, audit e receipts. |
| `members/61_member_model`, repository/aggregate derivation, authz e mutation plans | Shared Members domain | Binding, funções/vigências, capability/policy e epochs não podem virar decisões do adapter. |
| `activities/v2/100_act_contracts` a handlers/plans de calendário, tema, chamada e justificativa | Shared ACT domain | Portar funções puras e injetar clock/storage; WIP normativo requer integração própria. |
| `shared/core/20_firestore`, `composition/95_dev_boundary`, `96_dev_transport`, `legacy/*` | Adapter Apps Script existente | UrlFetch, PropertiesService, ScriptApp, envelope HTTP, fences de destino. |
| `workspace/*` e provider Drive/Gmail | Ports Workspace e adapters específicos | Tokens e efeitos externos continuam no backend; não fazer benchmark gerar comunicação real. |
| `sdk.cjs`, `sdk-host.cjs`, `wire.cjs`, `session.cjs` | Adapter Node experimental | SDK, transação, autenticação de infraestrutura e wire conversion; sem schema paralelo. |
| `worker.cjs`, `test-support/*` | Ponte/testes B0 | Não promover automaticamente o sandbox síncrono como servidor definitivo. |

Estrutura futura apenas proposta:

```text
geapa-core/
  estrutura Apps Script atual (ou apps-script/, se necessário)
  shared/domain/contracts/
  server/command-executor/
  prototype/node-command-executor/
```

Não foram movidas centenas de fontes para atender essa proposta. A composição
atual segue os paths de v43. B1 deve definir o executor único por operação,
identidade DEV e roteamento controlado; preservar commandRequests, transações,
idempotência, audit, fences e rebuild direcionado. Nenhum link Vercel é criado
por este documento. `prototype/` nunca é selecionado automaticamente como Root
Directory de produção.
