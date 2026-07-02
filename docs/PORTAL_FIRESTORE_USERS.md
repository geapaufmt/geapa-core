# Firestore portalUsers

`portalUsers/{uid}` e apenas cache operacional para acelerar a abertura do Portal GEAPA depois do Firebase Auth. A fonte normativa continua sendo GEAPA-CORE + PESSOAS v2.

## Snapshot seguro

`corePortalBuildFirestoreUserSnapshot(entrada, opts)` usa `corePortalResolverUsuarioAtual` e retorna somente os campos permitidos para abrir a interface:

- `uid`
- `idPessoa`
- `nomeExibicao`
- `email`
- `rga`
- `portalAtivo`
- `modoAcesso`
- `motivoBloqueio`
- `mensagemBloqueio`
- `perfilPortalEfetivo`
- `perfisPortal`
- `permissoes`
- `tipoVinculoAtual`
- `statusVinculoAtual`
- `cargoFuncaoAtual`
- `source: "GEAPA_CORE_PESSOAS_V2"`
- `sourceUpdatedAt`
- `cacheUpdatedAt`
- `cacheExpiresAt`
- `schemaVersion: "portal-user-v2"`

Nao incluir CPF, telefone, data de nascimento, observacoes internas, logs, pendencias detalhadas, chaves ou IDs de planilhas.

## Escrita no plano Spark

Como `portal-geapa` permanece no plano Spark, o CORE escreve por Firestore REST usando OAuth do Apps Script:

- `UrlFetchApp.fetch(...)`
- `Authorization: Bearer ScriptApp.getOAuthToken()`
- metodo `PATCH`
- documento `portalUsers/{uid}`

O transporte REST compartilhado fica em `24_core_firestore_rest.js`; o fluxo
de usuario apenas monta o snapshot seguro e delega encoding e escrita para essa
infraestrutura.

Nao ha Cloud Function, Secret Manager, service account ou chave privada nesse caminho.

Configure em Script Properties:

- `GEAPA_CORE_FIRESTORE_PROJECT_ID=portal-geapa`
- `GEAPA_CORE_FIRESTORE_DATABASE_ID=(default)` opcional

Se `GEAPA_CORE_FIRESTORE_PROJECT_ID` nao estiver configurado, as funcoes de sync retornam `FIRESTORE_PROJECT_ID_NAO_CONFIGURADO`.

O Apps Script precisa de autorizacao OAuth para Firestore/Datastore:

- `https://www.googleapis.com/auth/datastore`
- `https://www.googleapis.com/auth/script.external_request`

Se o projeto usar `oauthScopes` explicitos no `appsscript.json`, inclua esses escopos preservando os demais ja usados pelo CORE.

## Funcoes publicas

- `corePortalBuildFirestoreUserSnapshot(entrada, opts)`
- `corePortalGerarSnapshotFirestoreUsuario(entrada, opts)`
- `corePortalSincronizarUsuarioFirestore(entrada, opts)`
- `corePortalInvalidarCacheFirestoreUsuario(idPessoaOuEmail, opts)`
- `corePortalSyncFirestoreUserByEmail(email, opts)`
- `corePortalSyncFirestoreUserByIdPessoa(idPessoa, opts)`
- `corePortalSyncFirestoreUsersFromPessoasV2(opts)`
- `corePortalDiagnosticarFirestoreUsersDev(opts)`

`corePortalDiagnosticarFirestoreUsersDev()` e read-only e retorna somente totais
de membros autorizados, documentos `portalUsers`, cobertura conhecida e codigos
recentes de falha quando `PORTAL_LOG_ACESSOS` possui eventos persistidos. O
diagnostico nao retorna e-mails, UIDs ou `ID_PESSOA`. Como parte dos erros de
`portalLoginFirebase` existe apenas no Execution Log, o retorno informa quando
essa trilha nao esta disponivel para consulta estruturada.

Para escrever em `portalUsers/{uid}`, informe `uid` em `opts` ou use mapas `uidByEmail`/`uidByIdPessoa` no lote. O Core nao inventa UID a partir de e-mail ou `ID_PESSOA`.

A invalidacao tambem escreve apenas o snapshot minimo, com `portalAtivo=false` e `cacheExpiresAt` no passado. Acoes sensiveis continuam obrigatoriamente validadas pelo Apps Script/GEAPA-CORE.
