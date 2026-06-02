# Changelog

## 2026-06-02

- Removidas da API publica do core as funcoes temporarias da migracao inicial legado -> v2.
- Mantidas as APIs permanentes de contratos, auditoria e recalculo v2.
- Criadas APIs publicas de leitura para Pessoas v2 e Vigencias v2.
- Criado recalc controlado de `PESSOAS_RESUMO_OPERACIONAL`.
- Ampliado o recalc de `PESSOAS_RESUMO_OPERACIONAL` para usar Pessoas v2, Vigencias v2 e views de Atividades v2 quando disponiveis via Registry.
- Criado recalc controlado de `MEMBROS_DETALHES.SEMESTRE_ATUAL` em Pessoas v2 com base no RGA.
- Criado `coreDiagnosticarPessoasResumoOperacionalV2` como diagnostico somente leitura.
- Exposta comparacao legado/v2 como diagnostico somente leitura.
- Documentado o historico da migracao em `docs/MIGRACAO_V2_HISTORICO.md`.
- Marcado `29_core_domains_v2_migration.js` como arquivo historico interno, sem uso em producao.
