# Historico da migracao v2

Este documento registra a migracao inicial legado -> v2 dos dominios centrais `PESSOAS` e `VIGENCIAS`.

## Planilhas v2 criadas

### PESSOAS v2 - DEV

- ID: `1sa1CZTsqdDEWKWLd5uDAiM-Y59ko9FLZfABL0wc0HVM`
- Papel: base v2 do dominio Pessoas.

Abas criadas:

- `PESSOAS_BASE`
- `PESSOAS_IDENTIFICADORES`
- `MEMBROS_DETALHES`
- `COLABORADORES_ACADEMICOS`
- `PARTICIPANTES_EXTERNOS_DETALHES`
- `VINCULOS_GEAPA`
- `MEMBROS_EVENTOS_VINCULO`
- `PESSOAS_COMUNICACAO_CONSENTIMENTOS`
- `PORTAL_ACESSOS_EXCECOES`
- `PESSOAS_RESUMO_OPERACIONAL`

### VIGENCIAS v2 - DEV

- ID: `1M_KPFn7sRjZmQMtfoVOSDuSwlJqq9BLBQ-UYahcDQJw`
- Papel: base v2 do dominio Vigencias.

Abas criadas:

- `SEMESTRES`
- `PERIODOS`
- `DIRETORIAS`
- `SEMESTRES_DIRETORIA`
- `CARGOS_CONFIG`
- `VIGENCIAS_FUNCOES`
- `VIGENCIAS_RESUMO_ATUAL`

## Resumo da migracao

A migracao inicial criou as estruturas v2, transferiu dados legados revisados para as novas abas e permitiu auditoria manual posterior.

Principais decisoes registradas:

- `PESSOAS_BASE` concentra a identidade comum.
- `VINCULOS_GEAPA` representa estados institucionais como `MEMBRO_EFETIVO`, `MEMBRO_EM_ESPERA` e `EGRESSO`.
- `EX_MEMBRO` fica como alias legado; o contrato novo deve preferir `EGRESSO`.
- `MEMBROS_DETALHES` guarda campos especificos de membro.
- `PESSOAS_COMUNICACAO_CONSENTIMENTOS` e a fonte de comunicacao segmentada.
- `VIGENCIAS_FUNCOES` e a fonte normativa temporal de cargos/funcoes.
- `PESSOAS_RESUMO_OPERACIONAL` e `VIGENCIAS_RESUMO_ATUAL` sao caches calculados, nao fontes normativas.

## Funcoes temporarias usadas

As funcoes abaixo foram usadas para preparacao/migracao inicial e deixaram de ser API publica do core:

- `coreEnsurePessoasV2DevSheets`
- `coreEnsureVigenciasV2DevSheets`
- `coreEnsureCentralDomainsV2DevSheets`
- `coreDiagnosticarCentralDomainsV2DevSheets`
- `coreDiagnosticarMigracaoPessoasV2`
- `coreDiagnosticarMigracaoVigenciasV2`
- `coreMigrarPessoasV2`
- `coreMigrarVigenciasV2`
- `coreMigrarDominiosCentraisV2`
- `coreResetPessoasV2DevDestino`
- `coreResetVigenciasV2DevDestino`
- `coreResetAndMigrarVigenciasV2`

## Funcoes removidas/depreciadas

As exportacoes publicas e entradas em `GEAPA_CORE.domainsV2` foram removidas para as funcoes temporarias de setup, diagnostico de migracao, dry-run de migracao, migracao real e reset de destino.

O arquivo `29_core_domains_v2_migration.js` permanece apenas como arquivo historico interno nesta fase, com comentario de topo indicando que nao deve ser chamado em producao. Ele nao deve voltar a ser exportado sem uma decisao explicita.

## Funcoes permanentes mantidas

Continuam como parte operacional permanente:

- `coreGetDomainsV2Schemas`
- `coreGetDomainsV2ContractKeys`
- `coreAuditarPessoasV2`
- `coreAuditarVigenciasV2`
- `coreAuditarDominiosCentraisV2`
- `coreRecalcularVigenciasResumoAtualV2`

## Recuperacao de codigo antigo

Se for necessario consultar ou restaurar a migracao inicial, use o historico Git:

```bash
git log -- 29_core_domains_v2_migration.js
git show <commit>:29_core_domains_v2_migration.js
```

Nao execute novamente migracoes ou resets de destino sem uma decisao operacional explicita e uma nova fase documentada.
