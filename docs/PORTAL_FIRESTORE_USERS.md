# Firestore portalUsers

> Desde o piloto de 2026-08-21, o projeto e resolvido por ambiente explicito. DEV/HOMOLOG usam o projeto Firebase DEV; PROD usa projeto separado. Nao use namespaces.

`portalUsers/{uid}` e apenas cache operacional para acelerar a abertura do Portal GEAPA depois do Firebase Auth. A fonte normativa continua sendo GEAPA-CORE + PESSOAS v2.

## Provisionamento apos login Firebase

O provisionamento aceita tanto o e-mail canonico da PESSOAS v2 quanto um alias
resolvido oficialmente para a mesma `idPessoa`. No caso de alias, o Core repete
a resolucao pela identidade Firebase e so grava o documento quando ambas as
resolucoes apontam para a mesma pessoa. O documento usa o e-mail autenticado
pelo Firebase, pois ele e a identidade associada ao UID e utilizada na
validacao do cache privado pelo navegador.

Logs de acesso e diagnosticos devem receber `ambiente` explicitamente. Uma
execucao DEV nunca recorre a entrada PROD de `PORTAL_LOG_ACESSOS`; se a entrada
DEV nao existir no Registry, o registro persistido falha de forma controlada e
o evento permanece apenas no Execution Log.

## Snapshot seguro

`corePortalBuildFirestoreUserSnapshot(entrada, opts)` usa `corePortalResolverUsuarioAtual` e retorna somente os campos necessarios para autorizacao e abertura da interface:

- `uid`
- `idPessoa`
- `nomePublico` e o alias temporario `nomeExibicao`
- `email`
- `emailNormalizado`
- `perfilOperacional`
- `ativo`
- `podeAcessarPortal`
- `podeLerDadosPrivados`
- `roles`
- `permissions`
- `portalAtivo`
- `perfilPortalEfetivo`
- `perfisPortal`
- `permissoes`
- `source: "PESSOAS_V2"`
- `sourceSystem: "geapa-core"`
- `sourceUpdatedAt`
- `cacheUpdatedAt`
- `cacheExpiresAt`
- `lastLoginAt`
- `provisionedAt`
- `stale`
- `staleReason`
- `schemaVersion: "portal-user-v2"`

Nao incluir RGA, CPF, telefone, endereco, data de nascimento, observacoes internas, logs, frequencia, justificativas, tokens, chaves ou IDs de planilhas.

`portalAtivo`, `perfilPortalEfetivo`, `perfisPortal` e `permissoes` permanecem
temporariamente como aliases de compatibilidade. O contrato principal para
Rules e read models privados e `ativo + podeAcessarPortal + stale=false`.

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

- `GEAPA_CORE_FIRESTORE_DEV_PROJECT_ID=<projeto-dev>`
- `GEAPA_CORE_FIRESTORE_DEV_DATABASE_ID=(default)` opcional
- `GEAPA_CORE_FIRESTORE_PROD_PROJECT_ID=portal-geapa`
- `GEAPA_CORE_FIRESTORE_PROD_DATABASE_ID=(default)` opcional

Ambiente ausente/invalido, projeto ausente ou DEV/PROD iguais retornam erro. Nao existe fallback para a propriedade antiga nas funcoes migradas.

O Apps Script precisa de autorizacao OAuth para Firestore/Datastore:

- `https://www.googleapis.com/auth/datastore`
- `https://www.googleapis.com/auth/script.external_request`

Se o projeto usar `oauthScopes` explicitos no `appsscript.json`, inclua esses escopos preservando os demais ja usados pelo CORE.

## Funcoes publicas

- `corePortalBuildFirestoreUserSnapshot(entrada, opts)`
- `corePortalGerarSnapshotFirestoreUsuario(entrada, opts)`
- `corePortalSincronizarUsuarioFirestore(entrada, opts)`
- `corePortalProvisionarFirestoreUserAutenticado(firebaseIdentity, opts)`
- `corePortalMarcarFirestoreUserInativoPorUid(uid, opts)`
- `corePortalInvalidarCacheFirestoreUsuario(idPessoaOuEmail, opts)`
- `corePortalSyncFirestoreUserByEmail(email, opts)`
- `corePortalSyncFirestoreUserByIdPessoa(idPessoa, opts)`
- `corePortalSyncFirestoreUsersFromPessoasV2(opts)`
- `corePortalDiagnosticarFirestoreUsersDev(opts)`

`corePortalDiagnosticarFirestoreUsersDev()` e read-only e retorna totais de
autorizados, membros efetivos ativos, diretoria/secretaria/admin tecnico,
orientadores, documentos ativos/inativos, pendentes, orfaos, duplicidades e
erros recentes. E-mails ficam mascarados e UIDs truncados. `includeEmails:true`
so libera e-mail completo quando o ambiente efetivo nao e PROD.

Uma pessoa autorizada sem documento recebe o estado
`AGUARDANDO_PRIMEIRO_LOGIN_FIREBASE`. Isso nao e falha de cadastro: o UID so
existe depois que a pessoa autentica no Firebase. O Core nunca inventa UID nem
cria documento identificado por e-mail.

No login, o Portal valida o ID token pela Identity Toolkit, confere UID, e-mail,
audience, issuer e projeto, resolve a autorizacao oficial no Core e somente
entao chama `corePortalProvisionarFirestoreUserAutenticado` com
`identityVerified:true`. Essa funcao e uma fronteira interna entre Apps Scripts;
nao deve ser exposta como rota que aceite identidade declarada pelo navegador.

Para escrever em `portalUsers/{uid}`, informe `uid` em `opts` ou use mapas `uidByEmail`/`uidByIdPessoa` no lote. O Core nao inventa UID a partir de e-mail ou `ID_PESSOA`.

A invalidacao escreve apenas os campos de controle com `ativo=false`,
`podeAcessarPortal=false`, `stale=true` e cache vencido. Se o UID negado ainda
nao possui documento, nenhum documento e criado. Acoes sensiveis continuam
obrigatoriamente validadas pelo Apps Script/GEAPA-CORE.

## Homologacao manual

1. Execute `corePortalDiagnosticarFirestoreUsersDev({includeEmails:false})`.
2. Confirme que pendentes aparecem como `AGUARDANDO_PRIMEIRO_LOGIN_FIREBASE`.
3. Faca login com um usuario autorizado ainda pendente.
4. Execute novamente o diagnostico e confirme o aumento de `totalPortalUsersAtivos`.
5. Repita o login e confirme que nao ha duplicata e que `lastLoginAt` foi atualizado.
6. Teste uma conta sem autorizacao e confirme `USUARIO_NAO_AUTORIZADO`; se ja havia documento, ele deve ficar `stale=true`.
7. Confirme nas Rules que o navegador le somente `portalUsers/{request.auth.uid}` e nunca escreve.
