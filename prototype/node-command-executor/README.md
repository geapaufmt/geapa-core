# Node Command Executor — EXPERIMENTAL / B0 / NOT PRODUCTION

Protótipo CLI local. Não é serviço, endpoint, projeto Vercel ou destino de deploy.
O Portal publicado continua usando Apps Script. Nenhum arquivo desta pasta ativa
um executor remoto, altera IAM ou substitui o Core publicado.

## Resultado preservado

B0 mediu o SDK oficial contra Firestore DEV real, com dados sintéticos isolados:

| Comando | Apps Script, referência | Node mediana warm |
|---|---:|---:|
| EDIT_ACTIVITY | 25,386 s | 5,085 s |
| CREATE_ACTIVITY | 42,343 s | 4,651 s |
| THEME_SUBMIT | não medido | 4,855 s |
| ATTENDANCE_FINALIZE, 24 pessoas | não medido | 6,649 s |
| JUSTIFICATION_DECIDE | não medido | 4,707 s |

Uma primeira amostra + três warm por comando. A primeira edição levou 11,093 s;
as warm ficaram entre 5,003 e 5,205 s. Não há estimativa p95/cold start cloud.
Baseline anterior e fixtures não são o mesmo agregado/ator; não atribuir toda a
diferença exclusivamente ao SDK nem prometer esses números em produção.

Rebuilds warm separados: detalhe 2,267 s; lista com 55 atividades 4,406 s;
frequência de uma pessoa 2,348 s. Readback e rebuild não entram no tempo do
comando. Audit/receipt/fence entram no commit. Reads incluem revalidações e
documentos inexistentes; RPCs são chamadas observadas no SDK, não faturamento.

`evidence/summary.sanitized.json` conserva somente métricas; `checks.sanitized.json`
conserva contagens das 56 verificações B0 aprovadas e da limpeza. Nenhum token,
UID de conta de teste, cadastro, snapshot de fixture ou journal foi importado.
**B0.5 não repetiu o benchmark remoto.** As adaptações de portabilidade foram
verificadas offline; os tempos acima pertencem à campanha original.

## Fonte exata e dependências

Node usado: 24.19.0; `@google-cloud/firestore` 7.11.6; `firebase-admin` 13.10.0.
Versões, resoluções e integridades do lockfile original foram preservadas;
somente a indicação de Node 24 foi acrescentada aos metadados do pacote.

O Core v43 é reconstruído localmente de
`fac44ff26436e11d0b91424f0a7aa1081b151ff9`, nunca de HEAD. O manifesto valida
194 arquivos e o hash
`ad106b2f5eeba28cd694e9e9bd642530acc24a25e7b07796be499270afe4fd96`.
As cópias reconstruídas ficam em `base/`, ignorado pelo Git. Não há segunda cópia
permanente desses 194 arquivos na branch. Um clone raso precisa obter o commit
fixado antes da reconstrução.

`test-support/` contém somente código de doubles e geradores sintéticos que o
B0 antes importava de outros workspaces. Não contém dados reais nem fixtures
geradas. `source-provenance.json` registra as origens lógicas e os hashes de
entrada. Caminhos absolutos e UIDs das contas da campanha foram retirados.
O catálogo técnico de 53 comandos do gerador sintético antigo foi preservado
separadamente; com relógio fixo, os 67 documentos gerados são idênticos aos do B0
original. Isso evita depender do antigo workspace de testes ou importar snapshots.

## Teste local, sem credencial e sem rede Firestore

No diretório desta pasta:

```text
npm ci --ignore-scripts --no-audit --no-fund
npm test
```

O teste reconstrói a baseline do Git, gera a fixture apenas em memória, executa
edição/replay/rejeição no worker relocado e comprova zero chamada Firestore.
Também verifica sintaxe e bloqueios de ambiente/campanha encerrada.

## Reprodução remota futura — desabilitada por padrão

**Não executar na etapa B0.5.** A campanha `B0_NODE_POC_20261012` está encerrada e
é rejeitada antes de abrir credenciais. A reprodução exige autorização própria,
novo namespace e contas sintéticas previamente autorizadas. Não criar contas,
keys, grants ou IAM por conveniência do harness.

Configuração local, fora de Git:

- `B0_ALLOW_REMOTE=1`: autorização técnica explícita para uma nova campanha.
- `B0_NAMESPACE`: novo ID `B0_NODE_POC_...`, nunca o da campanha encerrada.
- `B0_ACTOR_UIDS`: UIDs sintéticos autorizados, separados por vírgula; BOARD,
  SELF e NONE devem corresponder ao provider privado.
- `B0_FIREBASE_TOOLS_ROOT`: diretório de uma instalação existente de
  `firebase-tools`, contendo `package.json` e `lib/`; alternativamente ele deve
  ser resolvível como dependência local. Nenhum cache OAuth é versionado.
- `B0_SESSION_PROVIDER`: caminho absoluto **fora do repositório** para o adapter
  privado já aprovado. Ele exporta `secrets()` com `{project, accounts}` e
  `authApi('signInWithPassword', body)` com `{value}`. Credenciais ficam nesse
  provider externo, nunca em argumentos de CLI, fixtures, fonte ou evidência.

O SDK continua exigindo o grant OAuth já ativo da conta institucional DEV
`geapadev@gmail.com`; `GOOGLE_APPLICATION_CREDENTIALS` e Emulator são rejeitados
no benchmark real. Nenhuma key dormente é ativada automaticamente.

Sequência reproduzível, somente após a autorização e configuração acima:

```text
node scripts/preflight.cjs
node benchmark.cjs
node security.cjs
node experiments.cjs create
node theme.cjs
node attendance.cjs
node justification.cjs
node security-extra.cjs
node projections.cjs
node closeout.cjs cleanup
node replay-parity.cjs
```

Journals e resultados dessa nova campanha ficam em `evidence/`, ignorados pelo
Git. Não apagar journals para forçar retry. `closeout.cjs` precisa da campanha
completa local, não apenas dos resumos sanitizados. A limpeza seleciona somente
paths registrados no namespace, usa preconditions e preserva receipts/audit.

## Fronteiras do protótipo

- `sdk.cjs`/`sdk-host.cjs`: SDK gRPC, `getAll`, queries limitadas, transações e
  preconditions. Projeto DEV e namespace B0 são obrigatórios.
- `worker.cjs`/`wire.cjs`: ponte local e representação dos contratos atuais;
  nenhum UrlFetch REST de rede é feito pelo Node.
- Firebase Admin verifica token e revogação. A lógica v43 continua validando
  binding, vínculo, vigência, capability, policy, escopo, revisão e estado.
- O provider privado e a lista explícita de contas são configuração do operador
  CLI; não vêm de um browser. Nenhum endpoint aceita esses parâmetros.
- `local.cjs`/`test-support/`: ambiente em memória e geradores de teste.

O worker síncrono é uma ponte de preservação, não a arquitetura definitiva de
B1. Veja `docs/NODE-B0-SHARED-BOUNDARIES.md` na raiz para a análise de extração
futura. Não há `server/command-executor`, integração Vercel ou deploy neste B0.5.
