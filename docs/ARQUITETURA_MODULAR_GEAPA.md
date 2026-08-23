# Arquitetura modular GEAPA

Documento de revisao arquitetural para redistribuicao gradual de responsabilidades entre os modulos do ecossistema GEAPA.

## Premissas

- Esta revisao nao altera planilhas, codigo operacional, exports publicos ou portal.
- `geapa-apresentacoes` deve ser tratado como legado/descontinuado para novas responsabilidades.
- Toda logica ativa de apresentacoes de membros deve convergir para `geapa-atividades`.
- `geapa-core` continua sendo biblioteca compartilhada, Registry, contratos, leitura por cabecalho, autorizacao base, diagnosticos e utilitarios.
- Modulos consumidores devem ser donos das regras de negocio de seus dominios.
- As bases v2 `PESSOAS` e `VIGENCIAS` ja foram migradas e conferidas manualmente. Elas nao devem ser tratadas como rascunho.
- Funcoes temporarias de migracao legado -> v2 nao devem voltar a ser usadas em fluxo operacional.

## Checkpoint Firestore de Atividades em DEV

O piloto de cadastro e agenda de Atividades foi concluido no projeto Firebase DEV, sem alteracao em PROD. O fluxo canonico desse recorte e:

```text
Portal DEV
  -> Firestore DEV (activities + activityPrivate)
  -> EXPORT_ATIVIDADES_FIRESTORE
```

- Firestore e a fonte da verdade de cadastro e agenda no DEV.
- `EXPORT_ATIVIDADES_FIRESTORE` e somente espelho/exportacao para consulta e compatibilidade operacional.
- Nao existe reverse sync de Sheets para Firestore nem dual-write bidirecional.
- `activityPrivate` e backend-only e nao deve ser exposta diretamente ao cliente web.
- Atividades normais nao possuem exclusao fisica operacional; cancelamento e ocultacao sao as transicoes suportadas.
- A aba legada `Atividades` nao volta a ser fonte canonica do recorte migrado.
- Presencas, justificativas, apresentacoes, envolvidos, convites, arquivos/materiais, notificacoes e logs ainda nao foram migrados para Firestore.
- O checkpoint vale apenas para DEV. Nao concede autorizacao para deploy, importacao ou write em PROD.

## Mapa de modulos

| Modulo | Estado | Papel oficial proposto |
| --- | --- | --- |
| `geapa-core` | Ativo | Infraestrutura compartilhada, Registry, contratos, APIs base, controle operacional, observabilidade, autorizacao portal e utilitarios comuns. |
| `geapa-membros-main-clean` | Ativo | Modulo operacional do dominio `PESSOAS`: ciclo de vida de membros, vinculos, eventos homologados, detalhes e resumo operacional. |
| `geapa-processo-seletivo` | Ativo | Dominio de seletivo: candidatos, etapas, comunicacoes do seletivo, aprovacao e entrega de eventos homologados ao dominio Pessoas. |
| `geapa-desligamentos-suspensoes` | Ativo | Dominio de desligamentos, suspensoes e retornos: pedidos, analise, homologacao e efeitos de saida/suspensao. |
| `geapa-atividades` | Ativo principal para atividades | Atividades internas, presencas, justificativas, motor disciplinar e apresentacoes integradas. |
| `geapa-comemoracoes` | Ativo | Rotinas comemorativas, aniversarios, mensagens especificas e consultas enxutas aos dados base. |
| `geapa-portal` | Ativo | Interface web e consumo de APIs. Nao deve conter regra critica de autorizacao sem validacao no backend. |
| Comunicacoes / Mensageria | Ativo, parcialmente no CORE | Hub de email, inbox/outbox, roteamento, templates base e contratos de mensagens. Regras especificas devem ficar nos modulos donos. |
| `geapa-apresentacoes` | Legado/descontinuado | Fonte historica de codigo e conhecimento. Nao deve receber novas responsabilidades nem novas dependencias. |

## Fronteiras gerais

O CORE deve decidir como acessar recursos e como validar contratos compartilhados. Ele nao deve decidir regras especificas de presenca, seletivo, desligamento ou apresentacao.

Separacao esperada:

- Resolucao de recursos: `Registry`.
- Controle operacional: `MODULOS_CONFIG`.
- Observabilidade operacional: `MODULOS_STATUS`.
- Identidade e leitura consolidada: APIs v2 de `PESSOAS` e `VIGENCIAS`.
- Regra de negocio: modulo dono do dominio.
- Interface final: portal.

Dependencias permitidas:

- Modulos ativos podem depender de `geapa-core`.
- `geapa-portal` pode depender de APIs publicas do CORE e dos modulos ativos publicados como backend.
- Camada de comunicacoes pode chamar adaptadores dos modulos ativos por contrato.

Dependencias proibidas ou a evitar:

- Modulo ativo novo depender de `geapa-apresentacoes`.
- CORE depender de `geapa-atividades`, `geapa-membros`, `geapa-processo-seletivo` ou outros modulos de negocio.
- Portal confiar apenas em permissao calculada no front-end.
- Modulo consumidor ler diretamente abas centrais quando existir API publica do CORE.

## Classificacao atual do GEAPA-CORE

### CORE_INFRAESTRUTURA

Arquivos/funcoes que devem permanecer no CORE como base compartilhada:

- `01_core_sheets.js`: normalizacao, header map, leitura/escrita por cabecalho, UX simples de abas.
- `02_core_dates.js`: datas e janelas de tempo.
- `03_core_email.js`: normalizacao de email e envio padrao.
- `04_core_gmail.js`: wrappers basicos de Gmail.
- `05_core_lock.js`: locks compartilhados.
- `06_core_log.js`: logs controlados.
- `07_core_assert.js`: validacoes genericas.
- `08_core_drive.js`: wrappers basicos de Drive.
- `09_core_http.js`: HTTP JSON generico.
- `10_core_registry.js`: Registry e resolucao por `KEY`.
- `11_core_sheet_records.gs`: leitura/escrita por registros e cabecalhos.
- `12_core_assets_registry.js`, `22_core_assets_service.js`: assets compartilhados.
- `13_core_governance.gs`: governanca institucional base.
- `23_core_modules_config.js`: controle operacional central.
- `24_core_modules_status.js`: status/log operacional central.
- `27_core_geapa_config.js`: `Config_GEAPA` vertical e compatibilidade temporaria com formato horizontal.
- `99_core_healthcheck.js`: diagnostico geral.

### CORE_CONTRATO_SCHEMA

Contratos, schemas e validacoes permanentes que podem ficar no CORE:

- `28_core_domains_v2_setup.js`: definicoes de schemas v2 devem permanecer como contrato, mas funcoes de criacao inicial de abas devem ficar arquivadas/depreciadas.
- `30_core_domains_v2_audit.js`: auditorias permanentes de `PESSOAS` e `VIGENCIAS`.
- `31_core_domains_v2_operational_api.js`: APIs de leitura v2 e recalculos operacionais permanentes.
- `docs/DOMINIOS_CENTRAIS_V2.md`.
- `docs/MIGRACAO_V2_HISTORICO.md`.

### CORE_AUTORIZACAO_PORTAL

Pode permanecer no CORE, pois autorizacao critica deve ser backend e compartilhada:

- `26_core_portal_access.js`.
- `corePortalResolverUsuarioAtual`.
- `corePortalCalcularPerfilEfetivo`.
- `corePortalListarPermissoesEfetivas`.
- `corePortalValidarAcesso`.
- `corePortalGetMeuResumo`.
- `corePortalDiagnosticarPerfisEPermissoes`.
- `corePrepararPortalParaV2`.

Observacao: funcoes que listam apresentacoes permitidas podem ficar temporariamente no CORE como ponte de autorizacao, mas o fornecimento operacional de apresentacoes deve ser do modulo `geapa-atividades`.

### CORE_API_BASE

APIs publicas base que devem permanecer como contrato minimo:

- `coreGetRegistry`, `coreGetRegistryRefByKey`, `coreGetRegistryMetaByKey`.
- `coreGetSheetByKey`.
- `coreReadRecordsByKey`, quando exportado ou consolidado.
- `coreAppendObjectByHeaders`, quando exportado ou consolidado.
- `coreWriteCellByHeader`.
- `coreGetGeapaConfigValue`, `coreGetGeapaConfigMap`, `coreGetGeapaConfigObject`.
- `coreGetModuleConfig`, `coreIsModuleEnabled`, `coreGetModuleMode`.
- `coreCanModuleUseCapability`, `coreAssertModuleExecutionAllowed`.
- `coreModuleStatusGet`, `coreModuleStatusMarkExecution`, `coreModuleStatusMarkSuccess`, `coreModuleStatusMarkError`, `coreModuleStatusMarkBlocked`.
- `corePessoasGetById`, `corePessoasFindByEmail`, `corePessoasFindByRga`.
- `corePessoasGetOperationalSummary`.
- `corePessoasListCurrentMembers`, `corePessoasListEffectiveMembers`, `corePessoasListExMembers`, `corePessoasListWaitingMembers`.
- `corePessoasListAcademicCollaborators`, `corePessoasListExternalParticipants`.
- `coreVigenciasGetCurrentFunctionByPessoa`, `coreVigenciasListCurrentFunctions`, `coreVigenciasGetPortalPermissionsByPessoa`.
- `coreSendEmailText`, `coreSendEmailHtml`, `coreSendTrackedEmail` ou equivalente oficial `coreEnviarEmailPadrao` se o nome for padronizado futuramente.

### REGRA_NEGOCIO_MEMBROS

Funcoes/arquivos hoje no CORE com forte cheiro de dominio `PESSOAS` e que devem ser avaliados para migrar ou reduzir:

- `16c_core_member_lifecycle_events.gs`: registro e atualizacao de eventos de ciclo de vida. O armazenamento generico pode ficar no CORE se for contrato compartilhado, mas processamento e efeitos devem ser do `geapa-membros-main-clean`.
- `16_core_member_identity_autofill.js`: autofill de identidade em abas de negocio deve ser consumido com cuidado. Parte generica pode ficar, mas regras de atualizacao de cadastro pertencem a `geapa-membros-main-clean`.
- `coreAppendMemberLifecycleEvent`, `coreListMemberLifecycleEvents`, `coreGetLatestMemberLifecycleEventByRga`, `coreUpdateMemberLifecycleEvent`, `coreUpdateMemberLifecycleEventStatus`.
- `coreSyncMembersCurrentDerivedFields`: compatibilidade legada com `MEMBERS_ATUAIS`; deve ser depreciada depois que v2 estiver operacional.
- `coreFindMemberIdentityByAny`, `coreFindMemberCurrentRowByAny`, `coreAutofillIdentityRowInSheet`: manter apenas se forem utilitarios genericos; caso contrario, migrar para membros.

Destino recomendado:

- `geapa-membros-main-clean` para processamento de eventos, criacao/atualizacao de vinculos, detalhes e resumo operacional.
- CORE mantem apenas APIs de leitura v2 e primitives genericas de escrita segura.

### REGRA_NEGOCIO_ATIVIDADES

Funcoes hoje no CORE com regra operacional de atividades devem ser migradas ou tratadas como compatibilidade:

- `geapaCoreListarMembrosParaChamada`.
- `coreListarMembrosParaChamada`.
- `listarMembrosParaChamada`.
- Testes fake relacionados a chamada em `80_testes.js`.
- Pontes de apresentacoes em `corePortalListarApresentacoesPermitidas` e `corePortalListarApresentacoesParaEgresso`, se elas passarem de autorizacao/filtro para regra operacional.

Destino recomendado:

- `geapa-atividades`, com consumo de `corePessoasListCurrentMembers`, `corePessoasGetOperationalSummary` e autorizacao via `corePortalValidarAcesso`.

### REGRA_NEGOCIO_SELETIVO

O CORE nao deve decidir aprovacao, etapas, listas de candidatos ou mensagens de seletivo. Pontos a observar:

- Adaptadores de mail com prefixo `SEL` ou regras fixas de seletivo no Mail Hub devem virar contrato/adaptador externo do `geapa-processo-seletivo`.
- CORE pode manter roteamento generico por `REGRAS_ROTEAMENTO`, mas nao semantica do processo seletivo.

Destino recomendado:

- `geapa-processo-seletivo`.

### REGRA_NEGOCIO_DESLIGAMENTOS

Pontos hoje no CORE:

- `coreMailCreateDesligamentosAdapter_` e correlacao `DES-*` em `18_core_mail_adapters.js`.
- Rotas e templates especificos de desligamento se existirem no Mail Hub.

Destino recomendado:

- `geapa-desligamentos-suspensoes`.
- CORE deve manter apenas contrato generico de adapter, fila e roteamento.

### REGRA_NEGOCIO_COMUNICACOES

O CORE pode manter a infraestrutura de comunicacoes, mas nao campanhas ou publicos especificos demais:

- `17_core_mail_hub.js`: manter como infraestrutura de inbox/outbox, eventos, anexos, config e regras de roteamento.
- `18_core_mail_adapters.js`: manter contrato de adapters, mas reduzir adapters hardcoded por modulo no longo prazo.
- `19_core_mail_renderer.js`: manter renderer base e templates institucionais comuns.
- `25_core_ex_members_recipients.js`: API de recipients de egressos pode ficar no CORE enquanto for leitura oficial de `PESSOAS`, mas campanhas e disparos devem ficar na Central de Mensageria ou modulo comunicacoes.

Destino recomendado:

- Um modulo/camada `COMUNICACOES` para campanhas, segmentacao final e execucao de disparos.
- CORE continua como fonte de leitura oficial e envio padrao.

### MIGRACAO_TEMPORARIA

Nao devem ficar em API publica nem ser chamadas sem decisao explicita:

- `29_core_domains_v2_migration.js`.
- `coreDiagnosticarMigracaoPessoasV2_`.
- `coreMigrarPessoasV2_`.
- `coreDiagnosticarMigracaoVigenciasV2_`.
- `coreMigrarVigenciasV2_`.
- `coreMigrarDominiosCentraisV2_`.
- `coreResetPessoasV2DevDestino_`.
- `coreResetVigenciasV2DevDestino_`.
- `coreResetAndMigrarVigenciasV2_`.
- Funcoes de criacao inicial em `28_core_domains_v2_setup.js`: `coreEnsurePessoasV2DevSheets_`, `coreEnsureVigenciasV2DevSheets_`, `coreEnsureCentralDomainsV2DevSheets_`, `coreDiagnosticarCentralDomainsV2DevSheets_`.

Status recomendado:

- Manter como arquivo historico interno ou remover apos validacao do portal v2.
- Nunca exportar novamente como API publica sem solicitacao explicita.

### DEPRECIAR

Candidatos a depreciacao apos validacao do portal v2 e dos modulos ativos:

- `geapaCoreBuscarMembroParaPortal`.
- `geapaCoreBuscarUsuarioPortal`.
- `geapaCoreBuscarMinhaSituacaoParaPortal`.
- `geapaCoreRunTesteUsuarioPortal`.
- `geapaCoreRunTesteMinhaSituacaoParaPortal`.
- `geapaCoreRunTesteListarMembrosParaChamada`.
- Funcoes baseadas em `MEMBERS_ATUAIS` como fonte principal.
- Exports ligados a `geapa-apresentacoes` ou prefixo `APR` quando existirem como regra ativa.

### MANTER_APENAS_COMPATIBILIDADE

Podem permanecer por uma fase curta para evitar quebra:

- Aliases `geapaCore*` usados por consumidores antigos.
- Compatibilidade horizontal de `Config_GEAPA`.
- Alias `EX_MEMBRO` para o contrato novo `EGRESSO`.
- `coreSyncMembersCurrentDerivedFields` enquanto ainda houver consumidores V1.
- Projecoes institucionais legadas de `CARGOS_INSTITUCIONAIS_CONFIG` enquanto `VIGENCIAS_RESUMO_ATUAL` assume o papel definitivo.

## Responsabilidades oficiais por modulo

### geapa-core

Responsavel por:

- Registry e resolucao de recursos.
- Leitura/escrita por cabecalho.
- Contratos e schemas centrais.
- APIs de leitura de `PESSOAS` v2 e `VIGENCIAS` v2.
- Recalculos de caches centrais quando eles forem parte do contrato compartilhado.
- Auditorias permanentes.
- Controle operacional por `MODULOS_CONFIG`.
- Status operacional por `MODULOS_STATUS`.
- Autorizacao base do portal.
- Mail Hub como infraestrutura.
- Envio padrao e utilitarios Gmail/Drive.

Nao responsavel por:

- Processar ciclo de vida de membros.
- Decidir presenca, falta, disciplina ou apresentacao.
- Decidir aprovacao de seletivo.
- Decidir desligamento/suspensao.
- Executar campanhas especificas.
- Renderizar experiencia do portal.

### geapa-membros-main-clean

Responsavel por:

- Dominio `PESSOAS`.
- Processar eventos homologados de ciclo de vida.
- Criar e atualizar `VINCULOS_GEAPA`.
- Manter `MEMBROS_DETALHES`.
- Atualizar ou solicitar recalculo de `PESSOAS_RESUMO_OPERACIONAL`.
- Refletir ingresso, espera, efetivacao, suspensao, retorno e egresso.
- Consumir APIs do CORE para escrita segura e leitura central.

Nao responsavel por:

- Frequencia, presenca e apresentacoes.
- Cargos/funcoes de vigencia.
- Regras de seletivo antes da homologacao.

### geapa-processo-seletivo

Responsavel por:

- Inscricoes, candidatos e etapas.
- Comunicacoes especificas do seletivo.
- Resultado final e homologacao de entrada.
- Entregar evento homologado para o dominio Pessoas.

Nao responsavel por:

- Criar regra institucional global.
- Atualizar planilhas centrais fora de contratos do CORE ou eventos homologados.

### geapa-desligamentos-suspensoes

Responsavel por:

- Pedidos de desligamento, suspensao e retorno.
- Analise administrativa do pedido.
- Homologacao e registro de evento.
- Acionar `geapa-membros-main-clean` ou API central para refletir estado consolidado.

Nao responsavel por:

- Recalcular presencas ou disciplina historica.
- Alterar vigencias sem evento/decisao oficial.

### geapa-atividades

Responsavel por:

- Atividades internas.
- Presencas.
- Justificativas de faltas.
- Motor disciplinar.
- Apresentacoes integradas.
- Arquivamento de periodos de atividades.
- Dados operacionais que substituem o uso ativo do antigo `geapa-apresentacoes`.

Nao responsavel por:

- Cadastro definitivo de pessoas.
- Vigencias institucionais.
- Regras de seletivo ou desligamento.

### geapa-comemoracoes

Responsavel por:

- Aniversarios e eventos comemorativos.
- Montagem de mensagens comemorativas.
- Consulta a dados de pessoas por APIs do CORE.

Nao responsavel por:

- Manter cadastro de pessoas.
- Decidir consentimento geral de comunicacoes abertas, salvo consumo de API oficial.

### geapa-portal

Responsavel por:

- Interface do usuario.
- Navegacao e experiencia.
- Consumo de APIs enxutas.
- Cache de sessao/tela quando apropriado.

Nao responsavel por:

- Autorizacao critica sem backend.
- Cruzamentos pesados que devem estar em views ou APIs consolidadas.

### Comunicacoes / Mensageria

Responsavel por:

- Disparos, campanhas, templates especificos e listas finais.
- Consumo de `coreGetExMembersCommunicationRecipients` e futuras APIs de recipients.
- Integracao com Mail Hub.

Nao responsavel por:

- Usar fila administrativa como base definitiva de destinatarios.
- Inferir estado de membro sem consultar `PESSOAS`/CORE.

## API minima recomendada do CORE

### Infraestrutura

- `coreGetSheetByKey(key)`.
- `coreGetRegistryRefByKey(key)`.
- `coreGetRegistryMetaByKey(key)`.
- `coreReadRecordsByKey(key, opts)`.
- `coreAppendObjectByHeaders(sheetOrKey, payload, opts)`.
- `coreWriteCellByHeader(sheet, rowNumber, headerMap, headerName, value, opts)`.
- `coreWithLock(fn, opts)`.
- `coreLogInfo`, `coreLogWarn`, `coreLogError`.

### Configuracao e status

- `coreGetGeapaConfigValue(key, opts)`.
- `coreGetModuleConfig(moduleName, flowName, opts)`.
- `coreIsModuleEnabled(moduleName, flowName, opts)`.
- `coreGetModuleMode(moduleName, flowName, opts)`.
- `coreCanModuleUseCapability(moduleName, flowName, capability, opts)`.
- `coreAssertModuleExecutionAllowed(moduleName, flowName, capability, opts)`.
- `coreModuleStatusMarkExecution(moduleName, flowName, capability, opts)`.
- `coreModuleStatusMarkSuccess(moduleName, flowName, capability, opts)`.
- `coreModuleStatusMarkError(moduleName, flowName, errorOrMessage, capability, opts)`.
- `coreModuleStatusMarkBlocked(moduleName, flowName, reasonCode, reasonMessage, capability, modeRead, opts)`.

### Pessoas e vigencias

- `corePessoasGetById(idPessoa)`.
- `corePessoasFindByEmail(email)`.
- `corePessoasFindByRga(rga)`.
- `corePessoasGetOperationalSummary(idPessoa)`.
- `corePessoasListCurrentMembers(opts)`.
- `corePessoasListEffectiveMembers(opts)`.
- `corePessoasListExMembers(opts)`.
- `corePessoasListWaitingMembers(opts)`.
- `corePessoasListAcademicCollaborators(opts)`.
- `corePessoasListExternalParticipants(opts)`.
- `corePessoasGetCurrentVinculo(idPessoa)` como API recomendada a consolidar, se ainda nao existir com esse nome.
- `coreVigenciasGetCurrentFunctionByPessoa(idPessoa)`.
- `coreVigenciasListCurrentFunctions(opts)`.
- `coreVigenciasGetPortalPermissionsByPessoa(idPessoa)`.

### Portal

- `corePortalResolverUsuarioAtual(email, opts)`.
- `corePortalValidarAcesso(idPessoa, permissaoOuPerfil, opts)`.
- `corePortalGetMeuResumo(email, opts)`.
- `corePortalListarPermissoesEfetivas(idPessoa, opts)`.

### Comunicacoes

- `coreSendEmailText(opts)`.
- `coreSendEmailHtml(opts)`.
- `coreSendTrackedEmail(params)`.
- Futuro alias institucional: `coreEnviarEmailPadrao(opts)`.
- `coreMailQueueOutgoing(contract)`.
- `coreMailProcessOutbox()`.
- `coreGetExMembersCommunicationRecipients(options)`.

## APIs que devem viver fora do CORE

### geapa-membros-main-clean

- `membrosProcessarEventoVinculo`.
- `membrosHomologarIngresso`.
- `membrosEfetivarMembroEmEspera`.
- `membrosRegistrarSuspensao`.
- `membrosRegistrarRetorno`.
- `membrosRegistrarEgresso`.
- `membrosRecalcularEstadoPessoa`.

### geapa-atividades

- `atividadesListarMembrosParaChamada`.
- `atividadesRegistrarPresenca`.
- `atividadesRegistrarJustificativaFalta`.
- `atividadesProcessarMotorDisciplinar`.
- `atividadesListarApresentacoesPermitidas`.
- `atividadesRegistrarApresentacao`.
- `atividadesArquivarPeriodo`.

### geapa-processo-seletivo

- `seletivoListarCandidatos`.
- `seletivoProcessarEtapa`.
- `seletivoEnviarComunicacaoCandidato`.
- `seletivoHomologarAprovado`.
- `seletivoGerarEventoIngresso`.

### geapa-desligamentos-suspensoes

- `desligamentosRegistrarPedido`.
- `desligamentosAnalisarPedido`.
- `desligamentosHomologarDesligamento`.
- `desligamentosHomologarSuspensao`.
- `desligamentosHomologarRetorno`.
- `desligamentosGerarEventoVinculo`.

### geapa-comemoracoes

- `comemoracoesListarAniversariantes`.
- `comemoracoesMontarMensagem`.
- `comemoracoesAgendarOuEnviar`.

### Comunicacoes

- `comunicacoesListarPublicoPorCampanha`.
- `comunicacoesEnviarCampanha`.
- `comunicacoesProcessarDescadastramento`.
- `comunicacoesRegistrarConsentimentoDeCampanha`.

## Plano para geapa-apresentacoes legado

`geapa-apresentacoes` nao deve receber novas funcoes, exports ou dependencias. O plano recomendado e:

1. Mapear funcoes uteis ainda nao migradas.
2. Classificar cada funcao como regra de atividade, helper generico ou legado descartavel.
3. Reimplementar regra ativa em `geapa-atividades`, consumindo CORE apenas para dados base.
4. Manter `geapa-apresentacoes` somente como fallback de emergencia durante uma janela curta.
5. Documentar que novas demandas de apresentacao pertencem a `geapa-atividades`.
6. Arquivar o repositorio ou remover triggers/deploys apenas depois de validacao operacional.

Funcoes/padroes candidatos a migracao:

- Ingestao e tratamento de emails de apresentacao.
- Convites e lembretes.
- Regras de titulo/eixo.
- Salvamento de anexos vinculados a apresentacoes.
- Relatorios historicos de apresentacoes.

Destino unico:

- `geapa-atividades`.

## Funcoes que devem sair da API publica do CORE

Remocao deve ser gradual, sem quebra imediata:

- Funcoes temporarias de migracao v2 ja listadas em `MIGRACAO_TEMPORARIA`.
- Wrappers `geapaCore*` legados apos portal v2 validado.
- Funcoes de chamada/presenca apos `geapa-atividades` expor API propria.
- Adapters hardcoded de modulos no Mail Hub, substituidos por registro de adapters pelos modulos.
- Funcoes de sync V1 de `MEMBERS_ATUAIS`, apos todos os consumidores usarem `PESSOAS` v2.
- Funcoes especificas de cargos legados baseadas em planilhas antigas, apos `VIGENCIAS` v2 virar fonte operacional completa.

## Estrategia gradual

### Fase 1 - Congelamento de fronteiras

- Documentar este contrato.
- Nao adicionar novas regras de negocio ao CORE.
- Nao criar novas funcoes em `geapa-apresentacoes`.
- Manter exports existentes para compatibilidade.

### Fase 2 - Publicar APIs minimas estaveis

- Consolidar nomes publicos que faltam, como `coreReadRecordsByKey`, `coreAppendObjectByHeaders` e `corePessoasGetCurrentVinculo`, se ainda nao estiverem exportados.
- Manter docs de consumo por modulo.
- Garantir que consumidores nao leiam planilhas centrais diretamente.

### Fase 3 - Mover regras de membros

- Levar processamento de ciclo de vida para `geapa-membros-main-clean`.
- CORE permanece como API de leitura/escrita segura e auditoria.
- Validar que `PESSOAS_RESUMO_OPERACIONAL` e recalculos continuam estaveis.

### Fase 4 - Mover atividades e apresentacoes

- Implementar APIs operacionais em `geapa-atividades`.
- Migrar logica util de `geapa-apresentacoes` para `geapa-atividades`.
- Trocar consumidores para `geapa-atividades`.
- Manter fallback legado por tempo determinado.

### Fase 5 - Modularizar comunicacoes

- Manter Mail Hub no CORE como infraestrutura.
- Registrar adapters por modulo ativo.
- Remover semantica hardcoded de `APR`, `SEL`, `DES` do CORE quando os modulos assumirem seus adapters.

### Fase 6 - Limpeza controlada

- Remover exports depreciados.
- Arquivar arquivos historicos de migracao.
- Arquivar `geapa-apresentacoes` apos validacao.
- Atualizar README e changelogs dos modulos.

## Riscos

- Quebra de consumidores antigos que ainda chamam aliases `geapaCore*`.
- Duplicacao temporaria de regra enquanto CORE e modulo ativo coexistem.
- Leitura direta de planilhas v2 por modulos, ignorando contratos do CORE.
- Confusao entre `EGRESSO` e alias legado `EX_MEMBRO`.
- Roteamento de email ainda preso a prefixos legados como `APR`.
- Portal consumir dados pesados se as views/API enxutas nao forem priorizadas.
- Apresentacoes ficarem divididas entre legado e `geapa-atividades` por tempo demais.

## Proximos passos recomendados

1. Validar este documento como fronteira oficial.
2. Criar checklist por modulo com funcoes a migrar.
3. Consolidar a API minima faltante no CORE sem remover compatibilidade.
4. Comecar pelo consumo operacional em `geapa-membros-main-clean` e `geapa-atividades`.
5. Migrar apresentacoes para `geapa-atividades` antes de qualquer limpeza pesada em `geapa-apresentacoes`.
6. So depois remover exports legados e arquivar codigo historico.
