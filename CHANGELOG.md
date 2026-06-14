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
- Criada camada de autorizacao v2 do Portal GEAPA baseada em Pessoas v2, Vigencias v2, `PORTAL_PERFIS` e `PORTAL_PERMISSOES`.
- Ajustado `corePortalResolverUsuarioAtual` para priorizar `PESSOAS_RESUMO_OPERACIONAL` e `VIGENCIAS_RESUMO_ATUAL` no login do Portal, mantendo `PORTAL_PERMISSOES` como fonte final de permissoes.
- Migrado `geapaCoreBuscarMinhaSituacaoParaPortal` para usar Pessoas v2 como fonte principal, mantendo `MEMBERS_ATUAIS` apenas como fallback.
- Criada API `corePortalGetOperationalConfig` para ler `PORTAL_CONFIG` com cache, normalizacao de tipos e filtro de chaves sensiveis.
- Criadas rotinas manuais seguras `corePessoasV2Diagnostico`, `corePessoasV2ConferirConsistencia`, `corePessoasV2AtualizarResumoOperacional`, `coreVigenciasV2Diagnostico`, `coreVigenciasV2ConferirConsistencia`, `coreVigenciasV2AtualizarResumoAtual` e `coreV2DiagnosticoGeral`.
- Criado `coreV2_jobDiarioManutencao` com instaladores manuais de trigger para orquestrar diagnostico, atualizacao e conferencia V2 sem ativacao automatica em producao.
- Criado `coreV2_runTesteJobDiarioDryRun` e documentada a checklist de homologacao manual dos jobs V2.
- Exportadas funcoes publicas para resolver usuario, perfil efetivo, permissoes efetivas, validacao de acesso, apresentacoes permitidas e diagnostico de prontidao do portal v2.
- Documentado o historico da migracao em `docs/MIGRACAO_V2_HISTORICO.md`.
- Marcado `29_core_domains_v2_migration.js` como arquivo historico interno, sem uso em producao.
