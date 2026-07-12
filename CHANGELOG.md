# Changelog

## 2026-07-12

- Criados contratos seguros para edicao de campos cadastrais de baixo risco do proprio usuario no Portal HOMOLOG.
- Criado fluxo separado de solicitacao, analise, aprovacao e aplicacao de correcoes sensiveis, com autorizacao `membros:analisar_correcoes`.
- Adicionado setup idempotente da aba `SOLICITACOES_ATUALIZACAO_CADASTRAL`, da key DEV correspondente e das permissoes DEV de Secretaria/Diretoria.
- Incluidos LockService, idempotencia, trilha de valores anterior/novo, mascaramento, deteccao de alteracao concorrente e recalculo de views.
- Adicionados 14 testes sem escrita real e checklist de publicacao exclusiva em HOMOLOG.
- Ajustado o fluxo para o projeto Apps Script unico: Registry DEV explicito, sem alternar a Script Property global `GEAPA_ENV`.
- Tornada opcional a fonte `PORTAL_PERMISSOES` DEV; sua ausencia nao bloqueia o setup nem provoca escrita na configuracao compartilhada/PROD.
- O setup real limpa o cache do Registry quando a key DEV ja foi cadastrada manualmente.

## 2026-07-11

- Criado `geapaCoreListarMembrosAdministracaoPortal` para a area administrativa somente leitura do Portal.
- A listagem revalida `membros:ler`, filtra apenas vinculos de membro, pagina resultados e usa lista branca de campos de `PESSOAS_V2_RESUMO_OPERACIONAL`.
- Incluidos filtros por identidade, vinculo, perfil, acesso ao Portal, pendencias e frequencia, sem expor dados cadastrais sensiveis.

## 2026-07-06

- Corrigido o recalculo central de `PESSOAS_RESUMO_OPERACIONAL` para conciliar identidade de `Atividades` com registros de `Atividades_Apresentacoes`.
- Incluidos frequencia consolidada, justificativas pendentes, mescla de intervalos de vinculo e diagnostico por campo.
- A escrita do cache passa a preservar colunas extras e valores manuais fora do contrato derivado.
- `CICLO_ULTIMA_APRESENTACAO` passa a ser o cabecalho canonico, mantendo leitura temporaria de `PERIODO_ULTIMA_APRESENTACAO`.
- Adicionado fallback explicito para leitura das fontes de Atividades v2 no ambiente DEV.
- Campos transitorios `PODE_*` em `CARGOS_CONFIG` deixam de bloquear o recalculo de `PESSOAS_RESUMO_OPERACIONAL` quando ausentes.
- `QTD_SEMESTRES_NO_GRUPO` passa a usar exclusivamente os semestres letivos de `VIGENCIAS_V2_SEMESTRES`, sem fallback para `CICLOS`.
- `CICLO_ULTIMA_APRESENTACAO` passa a gravar o ciclo real da ultima apresentacao, separado do alias legado de periodo/semestre.
- Criada API `geapaCoreBuscarMeuPerfilParaPortal` para expor, com escopo proprio e somente leitura, os dados cadastrais usados pela nova tela "Meu perfil" do Portal.

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
- Criado bootstrap seguro `coreV2_bootstrapConfiguracao`, conferencia `coreV2_conferirConfiguracao` e teste `coreV2_runTesteBootstrapDryRun` para preparar `MODULOS_CONFIG` e `MODULOS_STATUS` em DEV sem sobrescrever dados.
- Criado `coreV2_runTesteResolverRegistryV2` para diagnosticar keys Pessoas/Vigencias V2 no Registry bruto em DEV e ajustadas leituras V2 para resolver ambiente explicitamente.
- Exportadas funcoes publicas para resolver usuario, perfil efetivo, permissoes efetivas, validacao de acesso, apresentacoes permitidas e diagnostico de prontidao do portal v2.
- Documentado o historico da migracao em `docs/MIGRACAO_V2_HISTORICO.md`.
- Marcado `29_core_domains_v2_migration.js` como arquivo historico interno, sem uso em producao.
