# Firestore read models de Atividades

## Arquitetura

O Firestore e somente cache/read model do Portal. A fonte oficial continua
sendo Google Sheets V2, acessado e validado pelo Apps Script.

Antes deste pacote, o Core ja possuia uma integracao REST especifica para
`portalUsers/{uid}` em `26_core_portal_access.js`, usando OAuth do Apps Script.
O pacote 1A extrai essa infraestrutura para `24_core_firestore_rest.js` e
mantem os wrappers de `portalUsers` compativeis.

Nao sao usados Cloud Functions, service account, Secret Manager, chave privada
ou escrita pelo front-end.

## Configuracao

Script Properties:

- `GEAPA_CORE_FIRESTORE_PROJECT_ID`
- `GEAPA_CORE_FIRESTORE_DATABASE_ID`, opcional, com default `(default)`

O `appsscript.json` ja declara:

- `https://www.googleapis.com/auth/datastore`
- `https://www.googleapis.com/auth/script.external_request`

O principal que executa o Apps Script tambem precisa de permissao IAM para
gravar no banco Firestore. Respostas `401`/`403` devem ser tratadas como falha
de autorizacao/configuracao, sem fallback para credencial privada.

## API publica controlada

- `coreFirestoreSetDocument(path, data, options)`
- `coreFirestoreDeleteDocument(path, options)`
- `coreFirestoreBatchSetDocuments(items, options)`
- `coreFirestoreDiagnosticar(options)`

Todas as escritas usam `dryRun: true` por padrao. Escrita real exige
`dryRun: false` explicito. O encoder aceita string, boolean, numero, data,
array e objeto/mapa.

Exemplo somente leitura:

```javascript
coreFirestoreDiagnosticar()
```

Exemplo de previa sem escrita:

```javascript
coreFirestoreSetDocument('portalActivities/ATV-2026-1-0001', {
  idAtividade: 'ATV-2026-1-0001'
}, {
  dryRun: true
})
```

## Ownership

O Core e responsavel apenas pelo transporte REST, encoding e configuracao. O
contrato `portalActivities/{ID_ATIVIDADE}` e a selecao dos registros pertencem
ao `geapa-atividades`. O `geapa-portal` apenas le o cache e mantem fallback para
Apps Script.

Nenhum trigger e instalado por este pacote.
