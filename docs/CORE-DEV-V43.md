# Core DEV v43 — baseline de código

Esta branch preserva a versão 43 de **GEAPA Backend — DEV**, no projeto
`geapa-convergencia-dev`. Não é promoção para PROD nem autorização de deploy.

O manifesto `baselines/core-dev-v43.json` relaciona os 194 arquivos publicados,
seus nomes Apps Script, tipos e hashes SHA-256. A ordenação por nome de
`{name,type,source}`, serializada como JSON, produz:

`ad106b2f5eeba28cd694e9e9bd642530acc24a25e7b07796be499270afe4fd96`

A estrutura `legacy/`, `shared/`, `members/`, `activities/`, `editorial/`,
`workspace/`, `jobs/` e `composition/` reproduz a composição publicada. Os antigos
arquivos da raiz foram substituídos por essa composição; o histórico Git mantém
as versões anteriores. Documentos antigos descrevem estados históricos, não a
arquitetura canônica corrente.

Arquitetura corrente: Portal com Firebase Auth e leituras Firestore autorizadas;
Core único para comandos privilegiados; Firestore como autoridade canônica.
Readers/rotinas Sheets preservados no código legado não voltam a ser autoridade
de fluxos canônicos. Jobs e providers continuam sujeitos às configurações remotas;
versionar seu código não os ativa.

O vínculo `.clasp.json` herdado foi removido desta branch, e `.claspignore` exclui
tudo. Não executar clasp push, deploy Firebase/Hosting/Rules ou publicação de
versão. Configuração e credenciais remotas não fazem parte desta baseline.

Estado remoto observado: Core v43; Hosting `c7a4d525ca386dad`; Rules
`04511e52-4264-4565-9662-03dabbdc8ea0`. Commits posteriores desta branch preservam
WIP e POC e não correspondem automaticamente a uma versão homologada.
