# Resolucao definitiva de dominios V2 por ambiente

## Arquitetura

O Registry continua sendo a unica fonte de localizacao. Cada dominio possui uma
key de banco e duas entradas possiveis, uma `DEV` e uma `PROD`:

```text
ambiente validado -> *_V2_DB -> uma planilha -> aba canonica em codigo
```

O resolvedor nunca consulta a entrada do outro ambiente. `options.ambiente`
tem precedencia; na sua ausencia, usa `GEAPA_ENV`, que deve ser `DEV` ou `PROD`.
`HOMOLOG` e um canal de publicacao do Portal e usa as bases `DEV`; nao e valor
valido no Registry.

APIs internas do Core:

- `core_getDomainSpreadsheetRef_(domain, options)`;
- `core_openDomainSpreadsheet_(domain, options)`;
- `core_getDomainSheet_(domain, logicalSheet, options)`;
- `core_validateDomainRegistry_(domain, options)`.

Exports para outros projetos Apps Script:

- `coreGetDomainSpreadsheet`;
- `coreGetDomainSheet`;
- `coreReadDomainRecords`;
- `coreValidateDomainRegistry`;
- `coreValidateAllDomainRegistries`.

## Mapa canonico

| Dominio | DB key | Abas logicas e canonicas |
|---|---|---|
| PESSOAS | `PESSOAS_V2_DB` | `BASE=PESSOAS_BASE`; `IDENTIFICADORES=PESSOAS_IDENTIFICADORES`; `MEMBROS_DETALHES`; `SOLICITACOES_ATUALIZACAO=SOLICITACOES_ATUALIZACAO_CADASTRAL`; `LINKS_PERFIS=PESSOAS_LINKS_PERFIS`; `COLABORADORES=COLABORADORES_ACADEMICOS`; `EXTERNOS=PARTICIPANTES_EXTERNOS_DETALHES`; `VINCULOS=VINCULOS_GEAPA`; `EVENTOS=MEMBROS_EVENTOS_VINCULO`; `CONSENTIMENTOS=PESSOAS_COMUNICACAO_CONSENTIMENTOS`; `ACESSOS_EXCECOES=PORTAL_ACESSOS_EXCECOES`; `RESUMO=PESSOAS_RESUMO_OPERACIONAL` |
| VIGENCIAS | `VIGENCIAS_V2_DB` | `SEMESTRES`; `CICLOS`; `DIRETORIAS`; `SEMESTRES_DIRETORIA`; `CARGOS_CONFIG`; `FUNCOES=VIGENCIAS_FUNCOES`; `RESUMO=VIGENCIAS_RESUMO_ATUAL` |
| ATIVIDADES | `ATIVIDADES_V2_DB` | os 16 nomes de `ATIVIDADES_V2_SHEETS`, com ancora `ATIVIDADES=Atividades` |

## Leitura

1. Resolve a DB key no ambiente solicitado.
2. Abre a planilha uma vez por execucao.
3. Procura a aba canonica.
4. Apenas quando a DB key ou a aba estiver indisponivel, tenta a key especifica
   do mesmo ambiente.
5. O fallback gera `DOMAIN_SPECIFIC_KEY_FALLBACK` com dominio, aba logica,
   ambiente, origem e ID mascarado.

Se DB key e key especifica apontarem para IDs diferentes, a leitura pela aba
canonica da DB continua inequivoca e gera `DOMAIN_REGISTRY_DIVERGENCIA`. Se a
aba canonica estiver ausente, o fallback especifico e identificado no retorno
de resolucao e no log.

## Escrita

Escritas usam exclusivamente a DB key e a aba canonica. Sao bloqueadas quando:

- DB key esta ausente ou inativa;
- ha mais de uma linha ativa da key no mesmo ambiente;
- a aba canonica esta ausente;
- DB key e key especifica divergem;
- um `writeContext` tenta alternar dominio, ambiente ou planilha.

Dry-run pode ler fallback, mas nunca escreve. Nenhum ID de Pessoas, Vigencias
ou Atividades fica fixo no codigo operacional.

## Fallback temporario

As keys por aba permanecem no mapa do Core somente para leitura compativel e
diagnostico. Novos consumidores nao devem referencia-las. A remocao futura
deve ocorrer apenas depois de `coreValidateAllDomainRegistries` confirmar que
todas as abas canonicas existem e os consumidores publicados usam a DB key.

## Diagnostico

Execute separadamente para `DEV` e `PROD`:

```javascript
coreValidateAllDomainRegistries({ ambiente: 'DEV' });
coreValidateAllDomainRegistries({ ambiente: 'PROD' });
```

O resultado contem ambiente, ID mascarado, abas ausentes, divergencias,
duplicidades e keys especificas ainda ativas. O diagnostico e somente leitura.

## Checklist de implantacao

1. Fazer merge dos PRs sem publicar ainda.
2. No Registry, cadastrar exatamente uma linha ativa `DEV` e uma `PROD` para
   cada DB key.
3. Garantir que cada linha aponta para a planilha correta e para a aba ancora.
4. Criar ou renomear manualmente as abas para os nomes canonicos; nao executar
   setup de escrita antes de conferir o diagnostico.
5. Rodar os dois diagnosticos somente leitura e resolver duplicidades e
   divergencias.
6. Publicar nova versao da Library Core.
7. Atualizar e publicar geapa-atividades com a nova Library.
8. Publicar a versao HOMOLOG do Portal com `ambienteDadosV2: 'DEV'`.
9. Homologar leituras e escritas com dry-run e depois com dados de teste DEV.
10. Somente apos aceite, publicar versoes PROD com `ambienteDadosV2: 'PROD'`.

## Checklist de rollback

1. Reapontar os consumidores para a versao anterior da Library, sem alterar o
   Registry.
2. Reverter a versao do modulo Atividades.
3. Reverter o deployment do Portal para a versao anterior.
4. Preservar as DB keys e os dados escritos; nao apagar abas nem linhas.
5. Conferir logs por `DOMAIN_*` e executar novamente os diagnosticos somente
   leitura.

## Alteracoes manuais no Registry apos o merge

- adicionar/validar `PESSOAS_V2_DB` em `DEV` e `PROD`, com `SHEET_NAME=PESSOAS_BASE`;
- adicionar/validar `VIGENCIAS_V2_DB` em `DEV` e `PROD`, com `SHEET_NAME=SEMESTRES`;
- adicionar/validar `ATIVIDADES_V2_DB` em `DEV` e `PROD`, com `SHEET_NAME=Atividades`;
- desativar duplicidades no mesmo ambiente;
- manter keys especificas temporariamente ativas somente durante a janela de
  compatibilidade e fazer seus IDs coincidirem com a DB key correspondente.

Nenhuma dessas alteracoes e executada automaticamente pelo codigo ou pelos
testes deste repositorio.
