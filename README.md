# GEAPA Core (Apps Script Library)

Biblioteca compartilhada do ecossistema GEAPA. O `geapa-core` centraliza acesso a planilhas por Registry, normalizacao de dados, servicos de e-mail/Gmail, leitura tabular orientada a cabecalhos, governanca institucional e utilitarios reutilizados pelos demais modulos.

---

## Objetivo

O core existe para:

- resolver planilhas por `KEY` institucional via Registry;
- expor uma API publica estavel para os modulos consumidores;
- centralizar leitura/escrita por cabecalho e por registros;
- unificar envio de e-mails, replies e rastreamento de threads;
- projetar ocupantes atuais de ocupacoes institucionais;
- sincronizar campos derivados em `MEMBERS_ATUAIS`;
- oferecer utilitarios comuns de texto, datas, identidade e logs.

---

## Areas principais

### Registry e acesso a planilhas

- resolve `KEY -> { spreadsheetId, sheetName }` conforme ambiente;
- abre `Sheet` diretamente via `coreGetSheetByKey(key)`;
- mantem cache de registry por execucao.

Funcoes centrais:

- `coreGetRegistry()`
- `coreGetRegistryRefByKey(key)`
- `coreGetSheetByKey(key)`
- `coreGetRegistryMetaByKey(key)`
- `coreClearRegistryCache()`

### MODULOS_CONFIG e controle operacional

Camada central para decidir se um modulo/fluxo pode executar em determinado ambiente.

Esta camada nao substitui o Registry:

- Registry resolve recursos institucionais: `KEY -> spreadsheet/sheet/folder/etc`;
- `MODULOS_CONFIG` controla comportamento operacional: modulo, fluxo, modo, ambiente e capabilities.

A aba `MODULOS_CONFIG` fica na mesma planilha geral do Registry e usa os cabecalhos:

- `MODULO`
- `FLUXO`
- `ATIVO`
- `MODO`
- `AMBIENTE`
- `PERMITE_TRIGGER`
- `PERMITE_EMAIL`
- `PERMITE_INBOX`
- `PERMITE_SYNC`
- `PERMITE_DRIVE`
- `JANELA_MINUTOS`
- `ULTIMA_ALTERACAO`
- `ALTERADO_POR`
- `OBS`

Ordem de resolucao:

1. `MODULO + FLUXO + AMBIENTE`
2. `MODULO + GERAL + AMBIENTE`
3. fallback para `MODULO + FLUXO + PROD`, quando o ambiente atual nao for `PROD`
4. fallback para `MODULO + GERAL + PROD`, quando o ambiente atual nao for `PROD`

Semantica operacional:

- `ATIVO = NAO` bloqueia a execucao;
- `MODO = OFF` bloqueia a execucao;
- `MODO = MANUAL` bloqueia execucao automatica por trigger, mas permite execucao manual;
- `MODO = DRY_RUN` permite leitura, log e diagnostico; o consumidor deve evitar escrita destrutiva e envio real;
- capabilities validas: `TRIGGER`, `EMAIL`, `INBOX`, `SYNC`, `DRIVE`.

Funcoes publicas:

- `coreGetModuleConfig(moduleName, flowName, opts)`
- `coreIsModuleEnabled(moduleName, flowName, opts)`
- `coreGetModuleMode(moduleName, flowName, opts)`
- `coreCanModuleUseCapability(moduleName, flowName, capability, opts)`
- `coreAssertModuleExecutionAllowed(moduleName, flowName, capability, opts)`
- `coreGetModulesConfigDebug()`
- `coreClearModulesConfigCache()`
- `coreApplyModulesConfigSheetUx(opts)`

Exemplo em modulo consumidor:

```javascript
var decision = GEAPA_CORE.coreAssertModuleExecutionAllowed(
  'ATIVIDADES',
  'APRESENTACOES',
  'TRIGGER',
  { executionType: 'TRIGGER' }
);

if (decision.dryRun) {
  // executar apenas leitura, log e diagnostico
}
```

Observacoes desta etapa:

- o modelo e SOFT OFF: os triggers podem continuar instalados, mas a funcao sai cedo quando a configuracao bloquear;
- o core vira a fonte unica de decisao sobre `MODULOS_CONFIG`;
- nesta fase, os modulos consumidores ainda nao foram migrados para chamar essa API;
- a instalacao/remocao automatica de triggers nao e alterada.

UX operacional da aba:

- `coreApplyModulesConfigSheetUx()` aplica congelamento da linha 1, filtro, notas nos cabecalhos, cores por grupo, larguras de coluna, formatacao de data/numero e listas suspensas;
- listas suspensas restritivas: `ATIVO`, `MODO`, `AMBIENTE` e capabilities `PERMITE_*`;
- listas suspensas orientativas, aceitando novos valores quando necessario: `MODULO` e `FLUXO`;
- valores sugeridos de `MODULO`: `CORE`, `MEMBROS`, `SELETIVO`, `COMUNICACOES`, `ATIVIDADES`, `DESLIGAMENTOS`, `APRESENTACOES`;
- `EVENTOS` nao entra como modulo sugerido nesta fase.

### MODULOS_STATUS e observabilidade operacional

Camada central leve para registrar status operacional dos modulos por `MODULO + FLUXO`.

Separacao de responsabilidades:

- `MODULOS_CONFIG` decide se um modulo/fluxo pode executar;
- `MODULOS_STATUS` registra o que aconteceu: execucao, sucesso, erro ou bloqueio por config;
- Registry continua responsavel apenas por resolver recursos institucionais.

A aba `MODULOS_STATUS` fica na mesma planilha geral do Registry e usa os cabecalhos:

- `MODULO`
- `FLUXO`
- `ULTIMA_EXECUCAO`
- `ULTIMO_SUCESSO`
- `ULTIMO_ERRO`
- `MENSAGEM_ULTIMO_ERRO`
- `ULTIMO_BLOQUEIO_CONFIG`
- `MOTIVO_ULTIMO_BLOQUEIO`
- `ULTIMO_MODO_LIDO`
- `ULTIMA_CAPABILITY`
- `EXECUCOES_24H`
- `BLOQUEIOS_24H`
- `SUCESSOS_24H`
- `ERROS_24H`
- `OBS`

Funcoes publicas:

- `coreModuleStatusGet(moduleName, flowName, opts)`
- `coreModuleStatusEnsureRow(moduleName, flowName, opts)`
- `coreModuleStatusMarkExecution(moduleName, flowName, capability, opts)`
- `coreModuleStatusMarkSuccess(moduleName, flowName, capability, opts)`
- `coreModuleStatusMarkError(moduleName, flowName, errorOrMessage, capability, opts)`
- `coreModuleStatusMarkBlocked(moduleName, flowName, reasonCode, reasonMessage, capability, modeRead, opts)`
- `coreGetModulesStatusDebug()`

Exemplo de uso por modulo consumidor:

```javascript
GEAPA_CORE.coreModuleStatusMarkExecution('ATIVIDADES', 'PERIODO_VIGENTE', 'SYNC', {
  modeRead: 'ON'
});

try {
  // rotina operacional do modulo
  GEAPA_CORE.coreModuleStatusMarkSuccess('ATIVIDADES', 'PERIODO_VIGENTE', 'SYNC', {
    modeRead: 'ON'
  });
} catch (err) {
  GEAPA_CORE.coreModuleStatusMarkError('ATIVIDADES', 'PERIODO_VIGENTE', err, 'SYNC', {
    modeRead: 'ON'
  });
  throw err;
}
```

Exemplo de bloqueio por config:

```javascript
GEAPA_CORE.coreModuleStatusMarkBlocked(
  'APRESENTACOES',
  'GERAL',
  'MODO_OFF',
  'Fluxo bloqueado por MODULOS_CONFIG',
  'TRIGGER',
  'OFF'
);
```

Observacoes desta V1:

- se a linha `MODULO + FLUXO` nao existir, as funcoes de marcacao criam a linha automaticamente;
- `coreModuleStatusGet()` nao cria linha por padrao, mas aceita `opts.createIfMissing = true`;
- os contadores `EXECUCOES_24H`, `BLOQUEIOS_24H`, `SUCESSOS_24H` e `ERROS_24H` sao incrementais brutos nesta fase;
- ainda nao ha janela deslizante real de 24 horas;
- a escrita e feita por cabecalho, sem depender de indice fixo de coluna.

### Rotinas V2 de Pessoas e Vigencias

As rotinas manuais de diagnostico, conferencia e atualizacao segura de `PESSOAS_RESUMO_OPERACIONAL` e `VIGENCIAS_RESUMO_ATUAL` estao documentadas em [`docs/rotinas-v2-pessoas-vigencias.md`](docs/rotinas-v2-pessoas-vigencias.md).

Estas rotinas usam Registry, `dryRun` por padrao, `MODULOS_CONFIG` quando disponivel e `MODULOS_STATUS` em best effort. Consumidores externos devem usar uma versao/implantacao publicada do Core.

O job coordenado `coreV2_jobDiarioManutencao(options)` executa diagnostico, atualizacao e conferencia de Pessoas/Vigencias e pode receber callbacks do modulo Atividades. O trigger correspondente so e instalado por chamada manual a `coreV2InstalarTriggerJobDiario(options)`. A homologacao manual esta em [`docs/jobs-v2-manutencao.md`](docs/jobs-v2-manutencao.md).

O bootstrap seguro `coreV2_bootstrapConfiguracao(options)` confere e, quando explicitamente autorizado em DEV, cria apenas linhas ausentes em `MODULOS_CONFIG` e `MODULOS_STATUS`. A operacao esta documentada em [`docs/bootstrap-configuracao-v2.md`](docs/bootstrap-configuracao-v2.md).

O diagnostico `coreV2_runTesteResolverRegistryV2()` verifica, em modo somente leitura, se as keys de Pessoas/Vigencias V2 existem no Registry bruto em `AMBIENTE=DEV`, com IDs mascarados. Detalhes: [`docs/core-v2-registry-diagnostico.md`](docs/core-v2-registry-diagnostico.md).

### Sheets e records

Camada reutilizavel para leitura e escrita sem depender de colunas fixas.

Funcoes centrais:

- `coreNormalizeHeader(value)`
- `coreBuildHeaderIndexMap(headers, opts)`
- `coreFindHeaderIndex(headerMap, headerName, opts)`
- `coreGetCellByHeader(row, headerMap, headerName, opts)`
- `coreSetRowValueByHeader(row, headerMap, headerName, value, opts)`
- `coreWriteCellByHeader(sheet, rowNumber, headerMap, headerName, value, opts)`
- `coreFreezeHeaderRow(sheet, headerRow)`
- `coreEnsureFilter(sheet, headerRow, opts)`
- `coreApplyHeaderNotes(sheet, notesByHeader, headerRow)`
- `coreApplyHeaderColors(sheet, groups, headerRow, opts)`
- `coreApplyDropdownValidationByHeader(sheet, rulesByHeader, headerRow, opts)`
- `coreAppendObjectByHeaders(sheet, payload, opts)`
- `coreReadSheetRecords(sheet, opts)`
- `coreReadRecordsByKey(key, opts)`
- `coreFindFirstRecordByField(records, headerName, value, opts)`
- `coreFindFirstRecordByAnyField(records, headerNames, value, opts)`
- `coreGetNearestFilledValueUp(sheet, rowNumber, colNumber)`

Esses helpers tambem passam a sustentar a UX reaplicavel de planilhas dos modulos, como notas operacionais, filtros e listas suspensas por cabecalho.

### Datas e semestre

Funcoes para datas operacionais e leitura da vigencia de semestres.

Funcoes centrais:

- `coreNow()`
- `coreStartOfDay(date)`
- `coreAddDays(date, days)`
- `coreIsSameDay(d1, d2)`
- `coreInWindowDay(date, startInclusive, endExclusive)`
- `coreFormatDate(date, tz, pattern)`
- `coreGetSemesterForDate(refDate)`
- `coreGetSemesterIdForDate(refDate)`
- `coreGetCurrentSemester(refDate)`
- `coreGetLastCompletedSemester(refDate)`
- `coreParseEntrySemesterFromRga(rga)`
- `coreGetStudentCurrentSemesterFromRga(rga, refDate)`
- `coreGetCompletedGroupSemesterCountFromEntrySemester(entrySemesterShort, refDate)`

### Identidade institucional

O ecossistema GEAPA passa a adotar a seguinte regra de identificadores:

- membros continuam usando `RGA` como identificador oficial;
- professores usam `ID_PROFESSOR`;
- participantes externos usam `ID_PARTICIPANTE_EXTERNO`.

Regras de geracao:

- `ID_PROFESSOR` segue o formato `PROF-0001`, `PROF-0002`, ...;
- `ID_PARTICIPANTE_EXTERNO` segue o formato `EXT-0001`, `EXT-0002`, ...;
- novos IDs sao criados apenas quando a celula estiver vazia;
- IDs existentes nunca sao recalculados ou renumerados;
- para externos, o core verifica duplicidade por e-mail antes de criar um novo ID.

Funcoes publicas principais:

- `coreFillMissingProfessorIds()`
- `coreFillMissingExternalIds()`
- `coreEnsureProfessorIdForRow(rowNumber)`
- `coreEnsureExternalIdForRow(rowNumber)`
- `coreFindExternalByEmail(email)`
- `coreValidateExternalEmailDuplicates()`

Observacao operacional:

- as funcoes em lote saneiam bases ja existentes;
- as funcoes por linha permitem que modulos ou projetos consumidores garantam IDs automaticamente em novos registros.

### Config_GEAPA

A aba `Config_GEAPA` da PLANILHA GERAL passa a ser lida pelo core no formato vertical, uma configuracao por linha.

Colunas esperadas:

- `KEY`
- `VALOR`
- `TIPO`
- `GRUPO`
- `ATIVO`
- `DESCRICAO`
- `OBSERVACOES`

Funcoes publicas:

- `coreGetGeapaConfigValue(key, opts)`
- `coreGetGeapaConfigMap(opts)`
- `coreGetGeapaConfigObject(opts)`
- `coreDebugGeapaConfig(opts)`

Exemplo:

```javascript
var emailOficial = coreGetGeapaConfigValue('EMAIL_OFICIAL', {
  required: true
});
```

Regras da V1:

- a aba e resolvida pelo Registry, preferencialmente por `CONFIG_GEAPA`, mantendo compatibilidade com a chave legada `DADOS_OFICIAIS_GEAPA`;
- por padrao, apenas linhas com `ATIVO = SIM` entram no mapa;
- chaves sao normalizadas para caixa alta, sem acentos e com `_`, como `CURSO_MAE` e `LOCAL_PADRAO_REUNIOES`;
- `TIPO` preserva o valor como texto, mas aplica validacoes simples para `EMAIL`, `URL`, `DATA` e `COR_HEX`;
- o modelo horizontal antigo, com chaves na linha 1 e valores na linha 2, ainda e aceito temporariamente como fallback legado, mas esta depreciado.

### Dominios centrais v2 em DEV

O core trata `PESSOAS v2 - DEV` e `VIGENCIAS v2 - DEV` como contratos operacionais consumiveis em DEV. A migracao inicial e a conferencia manual ja foram concluidas; as funcoes permanentes agora devem ler, auditar e recalcular caches v2 sem refazer migracao legado -> v2.

Planilhas DEV preparadas:

- `PESSOAS v2 - DEV`
- `VIGENCIAS v2 - DEV`

Funcoes publicas:

- `coreGetDomainsV2Schemas()`
- `coreGetDomainsV2ContractKeys()`
- `coreAuditarPessoasV2()`
- `coreAuditarVigenciasV2()`
- `coreAuditarDominiosCentraisV2()`
- `coreCompararLegadoComV2(opts)`
- `coreRecalcularVigenciasResumoAtualV2(options)`
- `coreRecalcularPessoasResumoOperacionalV2(options)`
- `coreRecalcularMembrosDetalhesSemestreAtualV2(options)`
- `coreDiagnosticarPessoasResumoOperacionalV2(options)`
- `corePessoasGetById(idPessoa)`
- `corePessoasFindByEmail(email)`
- `corePessoasFindByRga(rga)`
- `corePessoasGetOperationalSummary(idPessoa)`
- `corePessoasListCurrentMembers(opts)`
- `corePessoasListExMembers(opts)`
- `corePessoasListWaitingMembers(opts)`
- `corePessoasListAcademicCollaborators(opts)`
- `corePessoasListExternalParticipants(opts)`
- `coreVigenciasGetCurrentFunctionByPessoa(idPessoa)`
- `coreVigenciasListCurrentFunctions(opts)`
- `coreVigenciasGetPortalPermissionsByPessoa(idPessoa)`

Comportamento permanente:

- expoe contratos/schemas v2;
- executa auditorias somente leitura;
- executa recalculos controlados de caches v2;
- expoe APIs publicas de leitura para Pessoas v2 e Vigencias v2;
- as funcoes temporarias de setup/migracao inicial foram removidas da API publica e documentadas em `docs/MIGRACAO_V2_HISTORICO.md`.

- as auditorias v2 sao somente leitura e retornam `ok`, `totalErros`, `totalAvisos`, `erros`, `avisos`, `recomendacoes` e `resumoQuantitativo`.
- `coreCompararLegadoComV2(opts)` e apenas diagnostico de cobertura entre fontes legadas e v2; nao altera dados e nao reabre a migracao.

Abas de `PESSOAS v2`:

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

Abas de `VIGENCIAS v2`:

- `SEMESTRES`
- `CICLOS`
- `DIRETORIAS`
- `SEMESTRES_DIRETORIA`
- `CARGOS_CONFIG`
- `VIGENCIAS_FUNCOES`
- `VIGENCIAS_RESUMO_ATUAL`

Observacoes normativas:

- `PESSOAS_RESUMO_OPERACIONAL` e cache/visao, nao fonte normativa;
- `COLABORADORES_ACADEMICOS` preserva `EIXO_ASSOCIADO` como dado cadastral academico, nao como fonte de destinatarios/divulgacao;
- `PESSOAS_COMUNICACAO_CONSENTIMENTOS` e a fonte de comunicacao segmentada; professores/tecnicos, externos e egressos devem registrar ali consentimento, status de comunicacao e interesses por eixo;
- `PESSOAS` nao passa a decidir presenca, falta, justificativa ou apresentacao;
- `VIGENCIAS_FUNCOES` registra funcao temporal e nao cria nova categoria de membro;
- `VIGENCIAS_RESUMO_ATUAL` e cache calculado: deve ser reconstruido por `coreRecalcularVigenciasResumoAtualV2`, nao editado manualmente como fonte normativa;
- permissao e perfil de portal devem continuar derivados de contratos oficiais, nao de edicao manual isolada em cache.

Leituras operacionais:

- buscas de pessoa retornam o cadastro base e, quando existir, identificadores, detalhes de membro, vinculos, resumo operacional, consentimentos e excecoes de portal;
- listas de membros atuais, egressos e membros em espera usam `VINCULOS_GEAPA` e, por padrao, apenas vinculos ativos;
- colaboradores academicos e participantes externos sao listados pelas abas de detalhe correspondentes, sem usa-las como base de comunicacao;
- funcoes vigentes usam `VIGENCIAS_FUNCOES` e agregam a configuracao correspondente em `CARGOS_CONFIG`;
- permissoes de portal por pessoa sao calculadas a partir de cargos vigentes e campos de permissao em `CARGOS_CONFIG`.

Recalculo seguro de `VIGENCIAS_RESUMO_ATUAL`:

```javascript
var previa = coreRecalcularVigenciasResumoAtualV2({
  dryRun: true
});
```

Se a amostra estiver correta e a auditoria estrutural estiver aceitavel:

```javascript
var resultado = coreRecalcularVigenciasResumoAtualV2({
  dryRun: false,
  confirmacao: 'RECALCULAR_RESUMO_ATUAL_V2'
});
```

O recalculo usa `VIGENCIAS_FUNCOES`, `CARGOS_CONFIG`, `PESSOAS_BASE` e `MEMBROS_DETALHES`; ele limpa apenas as linhas de dados do cache `VIGENCIAS_RESUMO_ATUAL` e reescreve o resumo atual por cabecalho.

Recalculo seguro de `PESSOAS_RESUMO_OPERACIONAL`:

```javascript
var previaPessoas = coreRecalcularPessoasResumoOperacionalV2({
  dryRun: true
});
```

Se a amostra estiver correta:

```javascript
var resultadoPessoas = coreRecalcularPessoasResumoOperacionalV2({
  dryRun: false,
  confirmacao: 'RECALCULAR_PESSOAS_RESUMO_V2'
});
```

O recalculo usa `PESSOAS_BASE`, `PESSOAS_IDENTIFICADORES`, `MEMBROS_DETALHES`, `VINCULOS_GEAPA`, `MEMBROS_EVENTOS_VINCULO`, `PORTAL_ACESSOS_EXCECOES`, `VIGENCIAS_RESUMO_ATUAL`, `SEMESTRES`/`CICLOS` e, quando disponiveis via Registry, as views de Atividades v2:

- `ATIVIDADES_V2_APRESENTACOES`
- `ATIVIDADES_V2_PORTAL_ATIVIDADES_DETALHES`
- `ATIVIDADES_V2_PRESENCAS_REGISTROS`
- `ATIVIDADES_V2_PORTAL_FREQUENCIA_MEMBROS`

`ATIVIDADES_V2_PORTAL_APRESENTACOES` nao e mais contrato ativo do Portal: apresentacoes publicas devem ser derivadas de `ATIVIDADES_V2_PORTAL_ATIVIDADES_DETALHES.APRESENTACOES_PUBLICAS_JSON`.

Campos recalculados:

- vinculo atual por prioridade operacional: membro efetivo ativo, membro em espera ativo, egresso, outros vinculos;
- cargo atual e perfil portal a partir de Vigencias, com fallback `Membro` para membro efetivo ativo sem funcao vigente;
- `PORTAL_ATIVO` a partir de vinculo, perfil e excecoes explicitas;
- tempo efetivo e quantidade de semestres considerando apenas vinculos `MEMBRO_EFETIVO`;
- apresentacoes realizadas, periodo da ultima apresentacao e frequencia resumida quando Atividades v2 tiver dados/views acessiveis;
- pendencias iniciais, flag de suspensao e elegibilidade basica para diretoria.

O recalc real faz upsert por `ID_PESSOA`: atualiza linhas existentes e adiciona ausentes, preservando cabecalhos e sem limpar a aba inteira. `DATA_LIMITE_ESTIMADA_DIRETORIA` fica vazia nesta V1 quando a regra de limite nao estiver suficientemente segura; o relatorio retorna `camposNaoCalculaveis` e um resumo `divergenciasLegado` somente leitura quando a comparacao for possivel.

Diagnostico somente leitura:

```javascript
var diagnostico = coreDiagnosticarPessoasResumoOperacionalV2();
```

O diagnostico verifica pessoas sem resumo, resumos sem pessoa, membro ativo sem portal/RGA/e-mail, egresso com portal ativo, divergencias contra `VIGENCIAS_RESUMO_ATUAL` e campos operacionais vazios.

Recalculo seguro de `MEMBROS_DETALHES.SEMESTRE_ATUAL`:

```javascript
var previaSemestreAtual = coreRecalcularMembrosDetalhesSemestreAtualV2({
  dryRun: true
});
```

Se a amostra estiver correta:

```javascript
var resultadoSemestreAtual = coreRecalcularMembrosDetalhesSemestreAtualV2({
  dryRun: false,
  confirmacao: 'RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2'
});
```

Esse recalc preenche somente `SEMESTRE_ATUAL` em `MEMBROS_DETALHES`, interpretando o ingresso no curso pelo RGA e o semestre vigente pela aba `SEMESTRES` de Vigencias v2. Ele nao altera `SEMESTRE_ENTRADA`, que representa o semestre de entrada do membro no GEAPA.

Historico da migracao inicial v2:

- a migracao legado -> v2 foi concluida e revisada manualmente;
- as funcoes temporarias de criacao, dry-run, migracao e reset foram removidas da API publica;
- o registro historico da fase fica em `docs/MIGRACAO_V2_HISTORICO.md`;
- o arquivo `29_core_domains_v2_migration.js` fica apenas como referencia interna temporaria e nao deve ser chamado em producao.

### Egressos para comunicacoes abertas

API oficial para consultar destinatarios egressos com consentimento ativo de comunicacao. O nome historico de funcao ainda usa `ExMembers` por compatibilidade temporaria.

Fonte de verdade:

- planilha `PESSOAS`;
- aba oficial legada `Ex-Membros`;
- nunca usar fila administrativa de pedidos como fonte de destinatarios futuros.

Funcao publica:

- `coreGetExMembersCommunicationRecipients(options)`

Funcao de diagnostico:

- `coreDebugExMembersCommunicationRecipients(options)`

Filtros obrigatorios aplicados pelo core:

- `EMAIL` preenchido e valido;
- `RECEBE_COMUNICACOES_GEAPA = SIM`;
- `STATUS_COMUNICACAO = ATIVO`;
- `STATUS_REGISTRO` com valor institucional valido, como `ATIVO`, `OK`, `VALIDO`, `REGULAR` ou `HOMOLOGADO`.

Filtro opcional por eixos:

```javascript
var recipients = coreGetExMembersCommunicationRecipients({
  eixos: ['EIXO_I', 'EIXO_III']
});
```

O filtro por eixos usa:

- preferencialmente `INTERESSE_EIXO_I` ate `INTERESSE_EIXO_VIII`;
- fallback/complemento em `EIXOS_INTERESSE`, aceitando lista separada por virgula, ponto e virgula, quebra de linha ou `|`.

Formato retornado:

```javascript
{
  nome: 'Nome da pessoa',
  rga: '202000000000',
  email: 'pessoa@email.com',
  eixosInteresse: ['EIXO_I', 'EIXO_III'],
  origem: 'EX_MEMBROS' // alias legado para publico de egressos
}
```

Regras de seguranca:

- o core deduplica por e-mail normalizado;
- este publico deve ser usado apenas para comunicacoes abertas, convites e divulgacoes compativeis com o consentimento;
- nao usar esta API para mensagens internas de membros ativos, comunicacoes administrativas restritas ou fluxos que dependam de vinculo atual.

### Eventos de ciclo de vida de membros

Camada compartilhada para registrar, listar, consultar o evento mais recente e atualizar eventos ja existentes em `MEMBER_EVENTOS_VINCULO` com escrita controlada pelo core.

Funcoes publicas:

- `coreAppendMemberLifecycleEvent(payload)`
- `coreListMemberLifecycleEvents(filters, opts)`
- `coreGetLatestMemberLifecycleEventByRga(rga, opts)`
- `coreUpdateMemberLifecycleEvent(eventId, patch)`
- `coreUpdateMemberLifecycleEventStatus(eventId, nextStatus, opts)`

Contrato atual de tipos e status:

- tipos suportados: `INGRESSO`, `DESLIGAMENTO_VOLUNTARIO`, `DESLIGAMENTO_POR_FALTAS`, `DESLIGAMENTO_ADMINISTRATIVO`, `SUSPENSAO`, `RETORNO`
- status suportados: `REGISTRADO`, `HOMOLOGADO`, `CANCELADO`, `PROCESSADO_ATIVIDADES`, `PROCESSADO_MEMBROS`

Campos atualizaveis via patch:

- `STATUS_EVENTO`
- `OBSERVACOES`
- `ATUALIZADO_EM`
- `PROCESSADO_POR_MODULO`
- `DATA_PROCESSAMENTO`
- `ERRO_PROCESSAMENTO`

Regras da API de update:

- localiza o registro por `ID_EVENTO_MEMBRO`
- valida o status contra o enum oficial do core
- rejeita campos fora da allowlist do contrato
- so altera colunas explicitamente permitidas
- e idempotente para retry: se o patch nao muda nada, o evento nao e regravado
- quando o patch muda algum campo e `ATUALIZADO_EM` nao e informado, o core grava automaticamente o timestamp da atualizacao

Observacao de schema:

- para manter compatibilidade com planilhas ja existentes, o core continua lendo o contrato atual sem exigir migracao previa;
- se `PROCESSADO_POR_MODULO`, `DATA_PROCESSAMENTO` ou `ERRO_PROCESSAMENTO` ainda nao existirem, o update pode autoestender a linha de cabecalho com essas colunas opcionais;
- se o modulo consumidor nao precisar desses metadados estruturados, pode continuar concentrando contexto tecnico em `OBSERVACOES`.

Exemplo de registro:

```javascript
GEAPA_CORE.coreAppendMemberLifecycleEvent({
  rga: '2023001',
  eventType: 'DESLIGAMENTO_POR_FALTAS',
  eventDate: new Date(),
  eventStatus: 'REGISTRADO',
  sourceModule: 'geapa-atividades',
  sourceKey: 'ATV-2026-001',
  notes: 'Evento aberto para homologacao.'
});
```

Exemplo de listagem:

```javascript
var events = GEAPA_CORE.coreListMemberLifecycleEvents({
  rga: '2023001',
  eventType: 'DESLIGAMENTO_POR_FALTAS'
}, {
  limit: 10
});
```

Exemplo de update generico:

```javascript
GEAPA_CORE.coreUpdateMemberLifecycleEvent('MEV-000001', {
  eventStatus: 'PROCESSADO_MEMBROS',
  observacoes: 'Desligamento efetivado no modulo de membros.',
  processedByModule: 'geapa-membros',
  processingDate: new Date(),
  processingError: ''
});
```

Exemplo de update focado em status:

```javascript
GEAPA_CORE.coreUpdateMemberLifecycleEventStatus('MEV-000001', 'HOMOLOGADO', {
  observacoes: 'Homologado pela gestao.'
});
```

### E-mail e Gmail

Camada compartilhada para envio HTML/texto, replies, labels e rastreamento.

Funcoes centrais:

- `coreIsValidEmail(email)`
- `coreNormalizeEmail(value)`
- `coreExtractEmailAddress(value)`
- `coreExtractDisplayName(value)`
- `coreUniqueEmails(values)`
- `coreSendEmailText(opts)`
- `coreSendEmailHtml(opts)`
- `coreSendHtmlEmail(opts)`
- `coreSendTrackedEmail(params)`
- `coreEnsureLabel(name)`
- `coreGetLabel(name)`
- `coreGetOrCreateLabel(name)`
- `coreSearchThreads(query, start, max)`
- `coreReplyThreadHtml(thread, subject, htmlBody, opts)`
- `coreMarkThread(thread, labelIn, labelOut)`

### Mail Renderer Institucional

Camada central para montar o HTML final dos e-mails do GEAPA sem tomar dos modulos o texto de negocio.

Responsabilidades do core nesta camada:

- cabecalho institucional;
- layout e estilos HTML;
- rodape e assinatura institucional;
- assunto final com `[GEAPA][CHAVE]`;
- draft compativel com a futura fila de saida.

Variantes disponiveis:

- `GEAPA_COMEMORATIVO`
- `GEAPA_OPERACIONAL`
- `GEAPA_CONVITE`
- `GEAPA_CLASSICO`

Funcoes publicas:

- `coreMailRenderEmailTemplate(templateKey, subjectHuman, payload)`
- `coreMailBuildFinalSubject(subjectHuman, correlationKey)`
- `coreMailBuildOutgoingDraft(contract)`

Contrato esperado do modulo para montar um draft:

- `moduleName`
- `templateKey`
- `correlationKey`
- `to`
- `cc`
- `bcc`
- `subjectHuman`
- `payload`

Estrutura esperada de `payload`:

- `title`, `subtitle`, `eyebrow`
- `introText` ou `introHtml`
- `blocks`: lista de blocos com `title`, `text`, `html`, `items` e `cta`
- `cta` opcional no nivel raiz
- `footerNote` e `preheader` opcionais

Observacao:

- os modulos continuam donos do conteudo de negocio;
- o renderer devolve `htmlBody`, `bodyText` e `emailOptions`, e agora tambem alimenta a V1 da `MAIL_SAIDA`;
- o lema exibido no rodape nao e fixo: ele e buscado da coluna `LEMA` da diretoria vigente em `VIGENCIA_DIRETORIAS`, com fallback temporario para a coluna legada `Slogan` e fallback seguro quando estiver vazio.
- a identidade oficial do grupo usada no renderer passa a ser lida pela camada central `Config_GEAPA`, incluindo nome oficial, sigla, e-mail oficial e cores institucionais;
- nesta etapa, `LOGO_OFICIAL` fica reservado para evolucao posterior; o renderer continua usando a imagem institucional padrao ja servida pelo core.
- `GEAPA_CLASSICO` preserva a linguagem mais simples do template historico de aniversarios: card unico, borda verde, lista linear de itens e rodape mais leve.

### Mail Hub (V1)

Camada central de ingestao de e-mails do Gmail para registro institucional em planilhas centrais.

Escopo atual:

- ler mensagens do Gmail com deduplicacao por `Id Mensagem Gmail`;
- extrair `Chave de Correlacao` do assunto no padrao `[GEAPA][CHAVE]`;
- resolver `Modulo Dono` e metadados de roteamento via `MAIL_REGRAS` e, quando nenhuma regra bater, via adapters por modulo;
- registrar eventos em `MAIL_EVENTOS`;
- manter upsert de indice em `MAIL_INDICE`;
- registrar metadados de anexos em `MAIL_ANEXOS`;
- consultar pendencias por modulo e marcar evento como processado;
- listar anexos operacionais por filtros reutilizaveis;
- reabrir o anexo real no Gmail sob demanda a partir de `MAIL_ANEXOS`;
- marcar o resultado operacional do anexo sem acoplar regra de negocio ao core.

Funcoes publicas:

- `coreMailRegisterModuleAdapter(adapter)`
- `coreMailGetModuleAdapter(moduleCodeOrName)`
- `coreMailListModuleAdapters()`
- `coreMailBuildCorrelationKey(moduleCodeOrName, ctx)`
- `coreMailParseCorrelationKey(key)`
- `coreMailResolveRouting(msgCtx)`
- `coreMailNormalizeOutgoingSubject(moduleCodeOrName, subject, ctx)`
- `coreMailRenderEmailTemplate(templateKey, subjectHuman, payload)`
- `coreMailBuildFinalSubject(subjectHuman, correlationKey)`
- `coreMailBuildOutgoingDraft(contract)`
- `coreMailQueueOutgoing(contract)`
- `coreMailProcessOutbox()`
- `coreMailIngestInbox(opts)`
- `coreMailGetConfig(key, defaultValue)`
- `coreMailGetConfigBoolean(key, defaultValue)`
- `coreMailGetConfigList(key)`
- `coreMailListPendingByModule(moduleName)`
- `coreMailGetLatestEvent(opts)`
- `coreMailListAttachments(opts)`
- `coreMailListPendingAttachments(opts)`
- `coreMailGetLatestPendingEventWithAttachment(opts)`
- `coreMailListAttachmentsByEvent(eventId, opts)`
- `coreMailGetAttachmentById(attachmentId, opts)`
- `coreMailGetAttachmentsByEvent(eventId, opts)`
- `coreMailMarkLatestPendingByModule(moduleName, processorName)`
- `coreMailMarkEventProcessed(eventId, processorName)`
- `coreMailMarkAttachmentProcessed(attachmentId, processorName, observations)`
- `coreMailMarkAttachmentSavedToDrive(attachmentId, processorName, driveInfo)`
- `coreMailMarkAttachmentIgnored(attachmentId, processorName, observations)`
- `coreMailMarkAttachmentError(attachmentId, processorName, observations)`
- `coreMailCleanupNoiseEvents()`
- `coreMailApplyOperationalSheetUx(opts)`

Arquitetura de adapters:

- o Mail Hub nao conhece regra de negocio de `APRESENTACOES`, `SELETIVO`, `MEMBROS` ou outros modulos;
- cada adapter declara apenas um contrato minimo comum: construcao de chave, parsing, match, resolucao de roteamento e normalizacao opcional de assunto;
- antes dos adapters, o core aplica regras ativas da aba `MAIL_REGRAS`, em ordem crescente de `Ordem`;
- o core faz registry dos adapters e escolhe o melhor roteamento com base em `correlationKey` ou heuristicas do proprio adapter quando nenhuma regra de planilha casar;
- adapters iniciais incluidos no core: `APR / APRESENTACOES`, `SEL / SELETIVO`, `MEM / MEMBROS`;
- o contrato permite formatos diferentes de chave entre modulos, desde que o proprio adapter saiba construir e interpretar a sua chave.

MAIL_REGRAS:

- permite roteamento configuravel sem alterar codigo para casos simples;
- cabecalhos esperados: `Id Regra`, `Ativa`, `Ordem`, `Campo Analise`, `Tipo Comparacao`, `Valor Comparacao`, `Modulo Dono`, `Tipo Entidade`, `Etapa Fluxo`, `Acao Quando Bater`, `Observacoes`, `Criado Em`, `Atualizado Em`;
- V1 operacional: `Acao Quando Bater = ROTEAR`;
- campos de analise suportados: `ASSUNTO`, `REMETENTE`, `DESTINATARIO`, `CORPO`, `TUDO`;
- comparacoes suportadas: `CONTEM`, `IGUAL`, `COMECA_COM`, `TERMINA_COM`, `REGEX`;
- exemplo: uma regra com `Campo Analise = ASSUNTO`, `Tipo Comparacao = CONTEM`, `Valor Comparacao = [GEAPA][APR-` e `Modulo Dono = APRESENTACOES` roteia replies de apresentacoes antes do fallback por adapter.

Observacao de plataforma:

- em Apps Script, registrar adapters com funcoes a partir de projetos consumidores via Library pode ter limitacoes de callbacks; a implementacao atual e segura para adapters definidos no proprio core ou no mesmo projeto.

Observacoes desta V1:

- nao migra os modulos consumidores existentes;
- nao implementa retry avancado da fila de saida;
- nao decide qual anexo pertence a qual regra de negocio;
- o roteamento por regras nesta V1 e propositalmente simples: escolhe a primeira regra ativa que bater e nao executa acoes complexas alem de `ROTEAR`.

Tratamento operacional de anexos (V1):

- o core continua registrando anexos em `MAIL_ANEXOS` no momento da ingestao;
- o modulo consumidor pode listar anexos pendentes por `moduleName`, `correlationKey`, `entityType`, `entityId`, `flowStep`, `statusAnexo`, `eventId`, `messageId`, `threadId` ou `attachmentId`;
- o core consegue reabrir a mensagem original no Gmail e devolver o anexo real sob demanda usando `Id Thread Gmail`, `Id Mensagem Gmail` e `Indice Anexo Mensagem`;
- o modulo consumidor continua responsavel por validar o arquivo, escolher o anexo correto e salvar no Drive;
- depois do processamento, o modulo pode marcar o anexo como `PROCESSADO`, `SALVO_DRIVE`, `IGNORADO` ou `ERRO`, incluindo `Processado Por`, `Data Hora Processamento`, `Observacoes`, `Id Arquivo Drive`, `Link Arquivo Drive` e `Pasta Destino Drive`;
- o indice central passa a considerar `SALVO_DRIVE` como anexo resolvido para a flag `Ha Anexo Pendente`.

MAIL_SAIDA (V1 minima):

- `coreMailQueueOutgoing(contract)` grava novas saidas com `Status Envio = PENDENTE`;
- `coreMailProcessOutbox()` processa a fila central, monta o assunto final com `[GEAPA][CHAVE]`, renderiza o HTML institucional, envia tecnicamente e atualiza `Id Thread Gmail`, `Id Mensagem Gmail`, `Enviado Em`, `Tentativas`, `Ultimo Erro` e `Status Envio`;
- ao enviar com sucesso, o core tambem registra `EMAIL_ENVIADO` em `MAIL_EVENTOS` e recompõe `MAIL_INDICE`;
- nesta V1, o contrato do modulo pode incluir `moduleName`, `templateKey`, `correlationKey`, `entityType`, `entityId`, `flowCode`, `stage`, `to`, `cc`, `bcc`, `subjectHuman`, `payload`, `priority`, `sendAfter` e `metadata`;
- quando o modulo optar por envio em massa via `bcc`, o core usa `EMAIL_OFICIAL` da camada central `Config_GEAPA` como envelope principal de seguranca.

Adapters institucionais da Mail Hub:

- `ATV` / `ATIVIDADES`
- `APR` / `APRESENTACOES`
- `SEL` / `SELETIVO`
- `MEM` / `MEMBROS`
- `DES` / `DESLIGAMENTOS`: usa chaves `DES-{ID_SOLICITACAO}`, roteia replies para `Modulo Dono = DESLIGAMENTOS` e preenche `Tipo Entidade = SOLICITACAO_VINCULO` / `Id Entidade = ID_SOLICITACAO`.

Higiene de ingestao e consistencia semantica:

- o hub le `MAIL_CONFIG` por chave e aplica regras de ingestao sem hardcode operacional;
- remetentes, dominios e assuntos tecnicos podem ser ignorados antes do registro;
- se `USAR_SOMENTE_ASSUNTOS_GEAPA = SIM`, o assunto precisa respeitar `ASSUNTO_PREFIXO_OBRIGATORIO`;
- replies com prefixos como `Re:` e `Fwd:` continuam aceitos quando o assunto base respeita o prefixo institucional;
- se o assunto ja trouxer uma `correlationKey` valida como `[GEAPA][CHAVE]`, a ingestao nao trata a mensagem como ruido por causa do prefixo;
- ruido tecnico de GitHub, Codex e alertas similares e preferencialmente descartado;
- se `MARCAR_RUIDO_COMO_IGNORADO = SIM`, o ruido passa a ser registrado com `Status Roteamento = IGNORADO` e `Status Processamento = IGNORADO`;
- anexos inline de replies, como logos citados pelo proprio Gmail, passam a ser ignorados na contagem de anexos recebidos;
- `coreMailApplyOperationalSheetUx(opts)` reaplica a UX operacional das abas centrais de e-mail, com notas nos cabecalhos, cores por grupo de coluna, filtros, congelamento da linha 1, validacoes por lista quando a aba permitir e visualizacao compacta nas colunas de texto longo;
- o indice e recomposto por resumo da chave de correlacao, preenchendo entidade, etapa, contadores e flags pendentes.

Schema minimo esperado na versao atual da planilha central:

- `MAIL_EVENTOS`: `Id Evento`, `Data Hora Evento`, `Direcao`, `Tipo Evento`, `Modulo Dono`, `Chave de Correlacao`, `Id Thread Gmail`, `Id Mensagem Gmail`, `Assunto`, `Email Remetente`, `Emails Destinatarios`, `Status Processamento`, `Processado Por`, `Data Hora Processamento`, `Possui Anexos`, `Quantidade Anexos`, `Criado Em`, `Atualizado Em`
- `MAIL_INDICE`: `Chave de Correlacao`, `Modulo Dono`, `Tipo Entidade`, `Id Entidade`, `Etapa Atual`, `Id Thread Gmail`, `Id Ultima Mensagem`, `Ultima Direcao`, `Ultimo Tipo Evento`, `Ultimo Email Remetente`, `Ultimo Assunto`, `Data Hora Ultimo Evento`, `Ha Entrada Pendente`, `Ha Anexo Pendente`, `Quantidade Eventos`, `Quantidade Entradas`, `Quantidade Saidas`, `Quantidade Anexos`, `Criado Em`, `Atualizado Em`
- `MAIL_ANEXOS`: `Id Anexo`, `Id Evento`, `Modulo Dono`, `Tipo Entidade`, `Id Entidade`, `Chave de Correlacao`, `Etapa Fluxo`, `Id Mensagem Gmail`, `Id Thread Gmail`, `Indice Anexo Mensagem`, `Nome Arquivo`, `Tipo Mime`, `Tamanho Bytes`, `Foi Salvo No Drive`, `Id Arquivo Drive`, `Link Arquivo Drive`, `Pasta Destino Drive`, `Status Anexo`, `Processado Por`, `Data Hora Processamento`, `Observacoes`, `Criado Em`, `Atualizado Em`
- `MAIL_REGRAS`: `Id Regra`, `Ativa`, `Ordem`, `Campo Analise`, `Tipo Comparacao`, `Valor Comparacao`, `Modulo Dono`, `Tipo Entidade`, `Etapa Fluxo`, `Acao Quando Bater`, `Observacoes`, `Criado Em`, `Atualizado Em`
- `MAIL_CONFIG`: `Chave`, `Valor`, `Ativo`
- `MAIL_SAIDA`: `Id Saida`, `Modulo Dono`, `Tipo Entidade`, `Id Entidade`, `Chave de Correlacao`, `Etapa Fluxo`, `Email Destinatario Principal`, `Emails Destinatarios`, `Emails Cc`, `Emails Cco`, `Nome Destinatario`, `Assunto`, `Corpo Texto`, `Corpo Html`, `Data Hora Agendada`, `Prioridade`, `Status Envio`, `Tentativas`, `Ultimo Erro`, `Id Thread Gmail`, `Id Mensagem Gmail`, `Enviado Em`, `Criado Em`, `Atualizado Em`, `Observacoes`

Configuracoes opcionais em `MAIL_CONFIG`:

- `GMAIL_QUERY_INGEST`
- `GMAIL_START`
- `GMAIL_MAX_THREADS`
- `GMAIL_MAX_MESSAGES_PER_THREAD`
- `ASSUNTO_PREFIXO_OBRIGATORIO`
- `USAR_SOMENTE_ASSUNTOS_GEAPA`
- `IGNORAR_REMETENTES`
- `IGNORAR_DOMINIOS`
- `IGNORAR_ASSUNTOS_REGEX`
- `MAX_EVENTOS_POR_EXECUCAO`
- `SALVAR_CORPO_COMPLETO`
- `MARCAR_RUIDO_COMO_IGNORADO`

Semantica pratica dessas configuracoes:

- `ASSUNTO_PREFIXO_OBRIGATORIO`: prefixo textual esperado no inicio do assunto, por exemplo `[GEAPA]`
- `USAR_SOMENTE_ASSUNTOS_GEAPA`: quando `SIM`, o hub ignora mensagens sem o prefixo obrigatorio
- `IGNORAR_REMETENTES`: lista por linha, virgula ou `;` de e-mails a ignorar
- `IGNORAR_DOMINIOS`: lista de dominios a ignorar, como `github.com`
- `IGNORAR_ASSUNTOS_REGEX`: lista de regex case-insensitive para assuntos tecnicos
- `MAX_EVENTOS_POR_EXECUCAO`: limite maximo de novos eventos registrados por execucao
- `SALVAR_CORPO_COMPLETO`: quando `SIM`, preenche `Corpo Texto`; quando `NAO`, salva apenas `Trecho Corpo`
- `MARCAR_RUIDO_COMO_IGNORADO`: quando `SIM`, registra ruido como `IGNORADO`; quando `NAO`, simplesmente nao registra
- `MAIL_ANEXOS` agora pode ser autoestendida pelo core quando os novos cabecalhos operacionais ainda nao existirem na planilha.

Testes manuais no projeto:

- `teste_getGeapaConfigValue_EMAIL_OFICIAL()`
- `teste_getGeapaConfigMap()`
- `diagnosticarConfigGeapa()`
- `test_core_domainsV2_auditarPessoas()`
- `test_core_domainsV2_auditarVigencias()`
- `test_core_domainsV2_auditarDominiosCentrais()`
- `test_core_domainsV2_compararLegadoComV2()`
- `test_core_domainsV2_recalcularVigenciasResumoAtual_dryRun()`
- `test_core_domainsV2_recalcularPessoasResumoOperacional_dryRun()`
- `test_core_domainsV2_recalcularMembrosDetalhesSemestreAtual_dryRun()`
- `test_core_domainsV2_diagnosticarPessoasResumoOperacional()`
- `test_core_domainsV2_recalcularPessoasResumoOperacional_REAL_CONFIRMADO()`
- `test_core_domainsV2_recalcularMembrosDetalhesSemestreAtual_REAL_CONFIRMADO()`
- `test_core_domainsV2_pessoasListCurrentMembers()`
- `test_core_domainsV2_recalcularVigenciasResumoAtual_REAL_CONFIRMADO()`
- `test_core_modulesConfig_debug()`
- `test_core_modulesConfig_clearCacheAndDebug()`
- `test_core_modulesConfig_applySheetUx()`
- `test_core_modulesConfig_atividades_geral()`
- `test_core_modulesConfig_apresentacoes_geral()`
- `test_core_modulesConfig_canTrigger_atividades()`
- `test_core_modulesConfig_assertTrigger_atividades()`
- `test_core_modulesConfig_canEmail_apresentacoes()`
- `test_core_modulesStatus_debug()`
- `test_core_modulesStatus_get_atividades_geral()`
- `test_core_modulesStatus_ensure_atividades_geral()`
- `test_core_modulesStatus_markExecution_atividades_geral()`
- `test_core_modulesStatus_markSuccess_atividades_geral()`
- `test_core_modulesStatus_markError_atividades_geral()`
- `test_core_modulesStatus_markBlocked_apresentacoes_geral()`
- `test_core_mailAdapters_list()`
- `test_core_mailAdapters_get_mem()`
- `test_core_mailAdapters_build_mem()`
- `test_core_mailAdapters_parse_mem()`
- `test_core_mailAdapters_resolveRouting_mem()`
- `test_core_mailRoutingRules_resolve_apresentacoes()`
- `test_core_mailRoutingRules_resolve_seletivo()`
- `test_core_mailAdapters_normalizeOutgoingSubject_mem()`
- `test_core_mailRenderer_render_operacional()`
- `test_core_mailRenderer_render_convite()`
- `test_core_mailRenderer_buildFinalSubject()`
- `test_core_mailRenderer_buildOutgoingDraft()`
- `test_core_mailOutbox_queue_operacional()`
- `test_core_mailOutbox_process()`
- `test_core_governance_currentBoardSlogan()`
- `test_core_mailHub_assertSchema()`
- `test_core_mailHub_config_read()`
- `test_core_mailHub_ingestInbox_dryRun()`
- `test_core_mailHub_ingestInbox_higiene_dryRun()`
- `test_core_mailHub_ingestInbox_real()`
- `test_core_mailHub_listPending_naoIdentificado()`
- `test_core_mailHub_listPending_membros()`
- `test_core_mailHub_getLatestEvent()`
- `test_core_mailHub_getLatestPending_membros()`
- `test_core_occupationCompat_headerAliases()`
- `test_core_occupationCompat_writePrefersOccupation_fakeSheet()`
- `test_core_memberLifecycle_updateEvent_patch_fakeSheet()`
- `test_core_memberLifecycle_updateEvent_invalidStatus_fakeSheet()`
- `test_core_portalBuscarMembroParaPortal_fakeSheet()`
- `test_core_portalBuscarMinhaSituacaoParaPortal_fakeSheet()`
- `test_core_portalV2_diagnosticarPerfisEPermissoes()`
- `test_core_portalV2_prepararPortalParaV2()`
- `test_core_portalV2_resolverUsuarioAtual_emailExemplo()`
- `test_core_mailHub_listPendingAttachments()`
- `test_core_mailHub_getLatestPendingEventWithAttachment()`
- `test_core_mailHub_getAttachmentById_example(attachmentId)`
- `test_core_mailHub_markAttachmentProcessed_example(attachmentId)`
- `test_core_mailHub_markAttachmentSavedToDrive_example(attachmentId)`
- `test_core_mailHub_markAttachmentError_example(attachmentId)`
- `test_core_mailHub_markLatestPending_membros_processed()`
- `test_core_mailHub_markLatestPending_naoIdentificado_processed()`
- `test_core_mailHub_cleanupNoiseEvents()`
- `test_core_mailHub_applyOperationalSheetUx()`

Sugestoes de melhoria na planilha:

- adicionar validacao de dados em `Direcao`, `Tipo Evento`, `Status Roteamento`, `Status Processamento`, `Status Conversa` e `Status Anexo` para evitar variacoes de texto;
- congelar a linha 1 em todas as abas e manter os nomes dos cabecalhos como estao hoje, porque o codigo depende deles por nome;
- na aba `Configuracoes`, preencher `GMAIL_QUERY_INGEST`, `GMAIL_MAX_THREADS` e `GMAIL_MAX_MESSAGES_PER_THREAD` para controlar a ingestao sem mexer no codigo;
- se quiser facilitar operacao humana, aplicar filtros nas abas `Eventos de Email`, `Indice de Conversas` e `Anexos`;
- se quiser ganhar rastreabilidade futura, vale considerar uma coluna opcional `Correlation Prefix` ou `Prefixo Correlacao` em `Indice de Conversas`, mas ela nao e necessaria para a V1 funcionar.

### Governanca institucional

Projecao de ocupantes atuais por ocupacao e por grupo de e-mails, com compatibilidade retroativa para bases que ainda usam `Cargo/Função` / `Cargo/Funcao`.

Funcoes centrais:

- `coreGetCurrentBoard(refDate)`
- `coreGetCurrentBoardSlogan(refDate)`
- `coreGetCurrentBoardMembers(refDate)`
- `coreGetCurrentBoardMembersByOccupation(occupation, refDate)`
- `coreGetCurrentBoardMemberByRole(role, refDate)`
- `coreGetCurrentBoardMemberByOccupation(occupation, refDate)`
- `coreGetCurrentLeadership(refDate)`
- `coreGetInstitutionalRolesActive()`
- `coreFindInstitutionalRoleByAnyName(text)`
- `coreGetCurrentInstitutionalAssignments(refDate)`
- `coreGetCurrentOccupantsByEmailGroup(groupName, refDate)`
- `coreGetCurrentContactsHtmlByEmailGroup(groupName, refDate)`
- `coreGetCurrentEmailsByEmailGroup(groupName, refDate)`
- `coreGetCurrentEmailsByRole(roleName, refDate)`
- `coreGetCurrentEmailsByOccupation(occupationName, refDate)`
- `coreSyncMembersCurrentInstitutionalOccupations(refDate)`

Compatibilidade semantica de ocupacao nesta etapa:

- o core aceita leitura de colunas `Ocupação`, `Ocupacao`, `Cargo/Função` e `Cargo/Funcao` para ocupacoes institucionais;
- para o campo atual em `MEMBERS_ATUAIS`, o core aceita `Ocupação atual`, `Ocupacao atual`, `Ocupação`, `Ocupacao`, `Cargo/Função atual` e aliases legados equivalentes;
- nas escritas, o core passa a preferir a coluna `Ocupação atual` quando ela existir, mantendo fallback automatico para os cabeçalhos legados;
- para renomeacoes institucionais catalogadas, o core resolve aliases historicos pelo `CARGOS_INSTITUCIONAIS_CONFIG`; nesta etapa, `Diretor(a) de Comunicação` passa a ser o nome principal e continua aceitando `Coordenador(a) de Comunicação`, `Coordenador de Comunicação` e `COORDENADOR_COMUNICACAO`;
- os nomes antigos de API com `Role` continuam funcionando para preservar compatibilidade, mas os aliases novos com `Occupation` passam a ser o caminho preferencial em integracoes novas.

### Identidade de membros

Busca e autofill com base em `MEMBERS_ATUAIS`.

Funcoes centrais:

- `coreNormalizeIdentityKey(value)`
- `coreFindMemberIdentityByAny(identity)`
- `coreFindMemberCurrentRowByAny(identity)`
- `coreAutofillIdentityRowInSheet(sheet, rowNumber, opts)`

Observacao:

- `coreAutofillIdentityRowInSheet` aceita `opts.nameHeaders`, `opts.rgaHeaders` e `opts.emailHeaders` para modulos com cabecalhos especificos, preservando o comportamento padrao quando `opts` nao e informado.

### Portal GEAPA

Contrato inicial entre `geapa-core` e `geapa-portal` para login por codigo e para a tela "Minha situacao".

Funcao publica exportada pela Library:

- `geapaCoreBuscarMembroParaPortal(emailOuRga)`
- `geapaCoreBuscarUsuarioPortal(emailOuRga)`
- `geapaCoreBuscarMinhaSituacaoParaPortal(emailOuRga)`
- `geapaCoreListarMembrosParaChamada(dataAtividade, contexto)`
- `geapaCoreRunTesteUsuarioPortal()`
- `geapaCoreRunTesteListarMembrosParaChamada()`
- `geapaCoreRunTesteMinhaSituacaoParaPortal()`

Regras do contrato:

- aceita e-mail ou RGA;
- normaliza entrada com `trim` e, para e-mail, `lowercase`;
- consulta `MEMBERS_ATUAIS` via Registry;
- a consulta cadastral retorna um unico membro ou `null`;
- a consulta de usuario retorna dados basicos seguros, cargos atuais do proprio usuario, perfis e permissoes iniciais;
- a consulta de "Minha situacao" retorna `ok: true` ou erro controlado com `ok: false`;
- nao retorna listas completas, dados sensiveis, frequencia detalhada, pendencias sensiveis, certificados ou historico;
- em caso de erro interno, nao expoe identificadores ou detalhes da planilha ao chamador.

Retorno em caso de sucesso:

```javascript
{
  id: string,
  nomeExibicao: string,
  emailCadastrado: string,
  rga: string,
  situacaoGeral: string,
  vinculo: string
}
```

Contrato de usuario autenticado:

```javascript
{
  ok: true,
  usuario: {
    id: string,
    nomeExibicao: string,
    rga: string,
    emailCadastrado: string,
    perfilPrincipal: "MEMBRO" | "DIRETORIA" | "PRESIDENCIA" | "SECRETARIA" | "COMUNICACAO" | "CONSELHO" | "ASSESSORIA",
    perfis: [
      "MEMBRO",
      "DIRETORIA",
      "PRESIDENCIA",
      "SECRETARIA",
      "COMUNICACAO",
      "CONSELHO",
      "ASSESSORIA"
    ],
    cargosAtuais: [
      {
        cargoKey: string,
        cargoNome: string,
        grupoCargo: string,
        fonte: "VIGENCIAS_DIRETORES" | "VIGENCIAS_ASSESSORES" | "VIGENCIAS_CONSELHEIROS",
        idDiretoria: string,
        dataInicio: "yyyy-MM-dd",
        dataFimPrevista: "yyyy-MM-dd"
      }
    ],
    permissoes: {
      podeVerAreaDiretoria: boolean,
      podeGerenciarAtividades: boolean,
      podeRegistrarChamada: boolean,
      podeEditarAtividade: boolean,
      podeAnalisarJustificativas: boolean,
      podeGerenciarCertificados: boolean,
      podeGerenciarComunicacao: boolean,
      podeGerenciarConfiguracoes: boolean
    }
  }
}
```

Regras de perfis do usuario:

- todo usuario autenticado recebe `MEMBRO`;
- cargo vigente em Diretores recebe `DIRETORIA`;
- `PRESIDENTE` e `VICE_PRESIDENTE` recebem `PRESIDENCIA`;
- `SECRETARIO_GERAL` e `SECRETARIO_EXECUTIVO` recebem `SECRETARIA`;
- `DIRETOR_COMUNICACAO` e `ASSESSOR_COMUNICACAO` recebem `COMUNICACAO`;
- `CONSELHEIRO_CONSULTIVO` recebe `CONSELHO`;
- cargos em Assessores recebem `ASSESSORIA` e perfis especificos quando o cargo indicar area;
- `ADMIN_TECNICO` nao e derivado automaticamente de vigencias.

Regras de permissoes iniciais:

- `MEMBRO` inicia com todas as permissoes falsas;
- `DIRETORIA` pode ver area da diretoria, gerenciar atividades, registrar chamada, editar atividade e analisar justificativas;
- `SECRETARIA` pode ver area da diretoria, gerenciar atividades, registrar chamada e analisar justificativas;
- `PRESIDENCIA` pode ver area da diretoria e gerenciar atividades;
- `COMUNICACAO` pode gerenciar comunicacao;
- `CONSELHO` fica sem acesso a area da diretoria nesta etapa, salvo se tambem tiver outro perfil que conceda essa permissao;
- nenhuma permissao critica deve ser validada somente no front-end.

Contrato de membros para chamada por atividade:

```javascript
geapaCoreListarMembrosParaChamada("2026-04-16", {
  perfil: "DIRETORIA",
  rga: "202311801000",
  email: "usuario@exemplo.com"
})
```

Retorno de sucesso:

```javascript
{
  ok: true,
  data: [
    {
      idPessoa: string,
      tipoParticipante: "MEMBRO",
      rga: string,
      email: string,
      nomeExibicao: string,
      situacao: "ATIVO" | "SUSPENSO" | "LICENCA" | "AFASTADO",
      vinculo: string,
      aplicavelNaData: true,
      contaPresenca: boolean,
      contaFalta: boolean,
      motivoNaoAplicavel: string
    }
  ],
  meta: {
    total: number,
    totalAplicaveis: number,
    totalNaoAplicaveis: number,
    dataReferencia: "yyyy-MM-dd",
    origem: "geapa-core",
    origemDados: "GEAPA_CORE",
    cacheHit: boolean,
    cacheTtlSeconds: 600,
    observacoes: []
  },
  performance: {
    totalMs: number,
    etapas: [
      { etapa: string, ms: number, totalMs: number }
    ]
  }
}
```

Regras da chamada:

- funcao somente leitura, segura para uso indireto pelo Portal via Apps Script/backend, como `geapa-atividades`;
- nao escreve em planilhas e nao altera producao;
- usa `MEMBERS_ATUAIS` como base principal de membros e, quando disponiveis via Registry, `PESSOAS_V2_BASE`, `PESSOAS_V2_MEMBROS_DETALHES` e `PESSOAS_V2_IDENTIFICADORES` para enriquecer identidade canonica;
- quando disponivel, usa `MEMBER_EVENTOS_VINCULO` para historico de ingresso, desligamento, suspensao e retorno;
- retorna apenas campos seguros para a chamada: `idPessoa`, `tipoParticipante`, `rga`, `email`, `nomeExibicao`, `situacao`, `vinculo`, `aplicavelNaData`, `contaPresenca`, `contaFalta` e `motivoNaoAplicavel`;
- a identidade canonica prioriza `ID_PESSOA`, depois `RGA`, depois e-mail, evitando duplicidade do mesmo membro real;
- nao retorna CPF, telefone, endereco, documentos, observacoes internas, dados disciplinares, logs ou IDs privados de planilhas;
- exclui pessoas que ingressaram depois da data da atividade ou que tiveram vinculo encerrado antes da data;
- quando situacao confiavel indicar suspensao, licenca ou afastamento, a pessoa pode retornar como N/A seguro, com `contaPresenca: false`, `contaFalta: false` e motivo padronizado, sem expor motivo sensivel;
- se o contexto trouxer perfil sem permissao operacional, retorna `PERMISSAO_NEGADA`;
- se `MEMBER_EVENTOS_VINCULO` ou campos oficiais de datas ainda nao estiverem disponiveis, a resposta inclui limitacoes em `meta.observacoes` e nao inventa dados ausentes.
- usa cache curto por ambiente/data com chave `CORE_MEMBROS_CHAMADA_V2:<AMBIENTE>:<yyyy-MM-dd>` e TTL de 10 minutos;
- o cache pode ser ignorado com `contexto.disableCache === true` e invalidado por data com `geapaCoreInvalidarCacheMembrosChamada(dataAtividade)`;
- retorna diagnostico de performance em `performance.totalMs` e `performance.etapas`.

Observacao operacional para o Pacote 3.1:

- `SALVAR` chamada em `geapa-atividades` e apenas rascunho operacional em `Portal_Acoes` (`CHAMADA_RASCUNHO_SALVO`) e nao deve ser tratado pelo Core como presenca oficial;
- apenas `FINALIZAR` produz presenca oficial em `Atividades_Presencas_Registros`;
- a listagem de membros do Core nao recalcula frequencia/justificativas e nao espera atualizacao dessas visoes apos `SALVAR`.

Erros controlados da chamada:

- data ausente ou invalida: `{ ok: false, errorCode: "DATA_ATIVIDADE_OBRIGATORIA", message: "Informe a data da atividade." }`
- perfil sem permissao: `{ ok: false, errorCode: "PERMISSAO_NEGADA", message: "Usuario sem permissao para listar membros para chamada." }`
- schema invalido: `{ ok: false, errorCode: "SCHEMA_MEMBROS_INVALIDO", message: "Base de membros sem cabecalhos obrigatorios para chamada." }`
- erro inesperado: `{ ok: false, errorCode: "ERRO_LISTAR_MEMBROS_CHAMADA", message: "Nao foi possivel listar membros para chamada." }`

### Portal GEAPA - perfis e permissoes

Camada inicial para ler autorizacao e permissoes do Portal GEAPA na planilha `PESSOAS`, sem Firebase Auth nesta etapa.

Abas oficiais esperadas via Registry:

- `PORTAL_PERFIS`
- `PORTAL_PERMISSOES`
- `PORTAL_CONFIG`
- `PORTAL_LOG_ACESSOS`

Abas de pessoas usadas para localizar e-mail:

- `Membros Atuais`
- `Membros em Espera`
- `Ex-Membros`

Colunas novas esperadas nas abas de pessoas:

- `PORTAL_ATIVO`
- `PERFIL_PORTAL`
- `PORTAL_OBS`

Funcoes publicas:

- `corePortalGetConfig()`
- `corePortalGetOperationalConfig(opts)` - le `PORTAL_CONFIG` com cache, filtra `ATIVO = SIM`, normaliza booleanos/numeros e oculta chaves sensiveis
- `corePortalClearConfigCache()`
- `corePortalGetProfiles()`
- `corePortalGetPermissionsByProfile(perfilPortal)`
- `corePortalAuthorizeEmail(email, opts)`
- `corePortalHasPermission(sessionOrEmail, permission, opts)`
- `corePortalLogAccess(payload)`
- `corePortalDiagnostics()`
- `corePortalResolverUsuarioAtual(entrada, opts)` - aceita e-mail, RGA, `ID_PESSOA` ou objeto com `email`, `rga`, `idPessoa`, `identificador`/`emailOuRga`
- `corePortalCalcularPerfilEfetivo(idPessoa, opts)`
- `corePortalListarPermissoesEfetivas(idPessoa, opts)`
- `corePortalValidarAcesso(idPessoa, permissaoOuPerfil, opts)`
- `corePortalGetMeuResumo(email, opts)`
- `corePortalListarApresentacoesPermitidas(email, options)`
- `corePortalListarApresentacoesParaEgresso(idPessoa, opts)`
- `corePortalDiagnosticarPerfisEPermissoes(opts)`
- `corePortalDiagnosticarAcessoPortalDev(opts)`
- `corePrepararPortalParaV2(opts)`

Regras principais:

- `corePortalAuthorizeEmail` e transitoria/legada enquanto nao ha Firebase Auth;
- a autorizacao por e-mail nao substitui autenticacao real e nao deve proteger acoes administrativas sensiveis sozinha;
- as funcoes v2 resolvem usuario, perfil e permissoes a partir de Pessoas v2, Vigencias v2 e `PORTAL_PERMISSOES`;
- `corePortalResolverUsuarioAtual` usa `PESSOAS_RESUMO_OPERACIONAL` como base principal de identidade/vinculo/perfil e `VIGENCIAS_RESUMO_ATUAL` como base principal de cargos atuais;
- `geapaCoreBuscarMembroParaPortal`, `geapaCoreBuscarUsuarioPortal` e `geapaCoreBuscarMinhaSituacaoParaPortal` usam Pessoas v2 como fonte principal e deixam `MEMBERS_ATUAIS` apenas como fallback de compatibilidade;
- `PORTAL_PERMISSOES` e a fonte oficial de permissoes efetivas;
- `CARGOS_CONFIG` define cargo e `PERFIL_PORTAL_PADRAO`, mas nao e fonte final de permissao;
- colunas `PODE_*` em `CARGOS_CONFIG` sao transitorias/depreciadas para autorizacao final;
- `ADMIN` deve vir de excecao explicita ativa em `PORTAL_ACESSOS_EXCECOES`, nao de cargo;
- `corePortalResolverUsuarioAtual` retorna a sessao canonica segura do Portal com `autenticado`, `idPessoa`, `perfilPortalEfetivo`, `perfisPortal`, `permissoes`, vinculo atual e cargos atuais sanitizados;
- `PORTAL_MODO_ACESSO` controla a liberacao geral: `TESTE` aplica `PORTAL_EMAILS_TESTE`; `MEMBROS_ATIVOS` ignora a whitelist de teste e libera membros ativos com `portal:acessar`; `PUBLICO_LIMITADO` so deve ser usado com fluxo publico seguro;
- para liberar aos membros hoje, configure `PORTAL_MODO_ACESSO = MEMBROS_ATIVOS`, `AUTH_ALLOW_VISITANTE = NAO` e `PORTAL_BLOCK_INACTIVE_MEMBERS = SIM`;
- apos alterar `PORTAL_MODO_ACESSO`, rode `corePortalClearConfigCache()` e force nova validacao de sessao no Portal;
- `EGRESSO` pode acessar apresentacoes ate a data de saida por `apresentacoes:ver_ate_saida`;
- `PORTAL_ATIVO = SIM` permite avaliar o perfil;
- `PORTAL_ATIVO = NAO` bloqueia;
- `PORTAL_ATIVO` vazio bloqueia por padrao, salvo configuracao explicita em `PORTAL_CONFIG`;
- `PERFIL_PORTAL` vazio usa `PORTAL_DEFAULT_PROFILE`, se existir, ou `MEMBRO` apenas para `Membros Atuais`;
- `Ex-Membros` bloqueiam por padrao;
- `Membros em Espera` so autorizam com `opts.includeWaiting === true` ou configuracao explicita;
- `VISITANTE` so e aceito com `AUTH_ALLOW_VISITANTE = SIM`;
- `ADMIN` nao e usado como padrao; precisa estar explicitamente em `PERFIL_PORTAL`;
- permissoes sempre vem de `PORTAL_PERMISSOES`, nao de codigo hardcoded.

Logs de acesso:

- `corePortalLogAccess(payload)` escreve somente em `PORTAL_LOG_ACESSOS`;
- nunca registre senha, token, segredo ou ID Token;
- campos usados: `TIMESTAMP`, `EMAIL`, `UID_FIREBASE`, `NOME`, `PERFIL_PORTAL`, `ACAO`, `RESULTADO`, `MOTIVO`, `ORIGEM`, `USER_AGENT`, `OBS`.

Firebase futuro:

- Firebase Hosting/Auth, validacao de ID Token, API key e service account ficam fora desta etapa;
- em PR futuro, a autenticacao Firebase deve chamar a mesma camada de perfis/permissoes, trocando a entrada `corePortalAuthorizeEmail(email)` por uma entrada autenticada como `corePortalAuthorizeFirebaseUser(idToken)`.

Diagnostico v2:

- `corePortalDiagnosticarPerfisEPermissoes()` verifica perfis obrigatorios, permissoes por perfil, usuarios sem perfil, egressos sem perfil `EGRESSO` e excecoes `ADMIN`;
- `corePortalDiagnosticarAcessoPortalDev()` verifica o modo de acesso, membros ativos com/sem e-mail, perfis sem `portal:acessar` e e-mails que ficariam bloqueados por `PORTAL_EMAILS_TESTE` no modo `TESTE`;
- `corePrepararPortalParaV2()` retorna `PRONTO`, `PARCIAL` ou `BLOQUEADO` antes de qualquer alteracao no `geapa-portal`;
- detalhes da arquitetura ficam em `docs/PORTAL_AUTORIZACAO_V2.md`.

### Portal GEAPA - conteudo publico editorial

`PORTAL_CONTEUDO_PUBLICO` e a planilha CMS editorial do Portal GEAPA. Ela serve para conteudo publico editavel em Google Sheets, como home, sobre, historia, parceiros, documentos, midias e complementos publicos de pessoas/gestoes.

Ela nao e fonte oficial para atividades, apresentacoes, membros, diretoria, frequencia ou permissoes. Esses dados continuam vindo dos modulos de dominio e de views/contratos `PORTAL_*` especificos.

Funcoes publicas:

- `corePortalPublicContentGetDefinitions()`
- `corePortalPublicContentEnsureStructure(options)`
- `corePortalPublicContentCreateSpreadsheet(options)`
- `corePortalPublicContentEnsureSheets(options)`
- `corePortalPublicContentEnsureHeaders(options)`
- `corePortalPublicContentDiagnostics(options)`
- `corePortalPublicContentReadRows(key, options)`
- `corePortalPublicContentGetPage(slug, options)`
- `corePortalPublicContentGetHome(options)`
- `corePortalPublicContentGetSobre(options)`
- `corePortalPublicContentGetHistoria(options)`
- `corePortalPublicContentGetParceiros(options)`
- `corePortalPublicContentGetDocumentos(options)`
- `corePortalPublicContentGetConfig(options)`
- `corePortalPublicContentGetMidias(options)`
- `corePortalPublicContentGetDiretoriaComplementos(options)`
- `corePortalPublicContentGetPessoasComplementos(options)`
- `corePortalPublicContentGetGestoesComplementos(options)`
- `corePortalPublicContentGetPessoasConfig(options)`
- `corePortalPublicContentBuildPublicSnapshot(options)`

Abas editoriais garantidas:

- `PUBLIC_HOME`
- `PUBLIC_SOBRE`
- `PUBLIC_HISTORIA`
- `PUBLIC_PARCEIROS`
- `PUBLIC_DOCUMENTOS`
- `PUBLIC_CONFIG`
- `PUBLIC_MIDIAS`
- `PUBLIC_PESSOAS_COMPLEMENTOS`
- `PUBLIC_GESTOES_COMPLEMENTOS`
- `PUBLIC_PESSOAS_CONFIG`
- `PUBLIC_LOG_PUBLICACAO`

`PUBLIC_DIRETORIA_COMPLEMENTOS` e `PORTAL_PUBLIC_DIRETORIA_COMPLEMENTOS` ficam apenas como legado defensivo de leitura. A rotina estrutural nova nao cria essa aba antiga.

Regras da rotina estrutural:

- usa Registry para resolver a planilha quando as keys ja existem;
- se a aba nao existir, cria a aba;
- se a aba existir, preserva dados e colunas extras;
- se faltarem colunas, adiciona ao final;
- rodar duas vezes nao duplica abas nem cabecalhos;
- usa `LockService` via helper do Core para evitar concorrencia;
- nao escreve no Registry automaticamente nesta etapa;
- `corePortalPublicContentCreateSpreadsheet({ confirm: true })` existe para criacao administrativa explicita e retorna linhas sugeridas para cadastro manual no Registry.

Firestore futuro:

- o Firestore sera espelho publico futuro, nao fonte editorial principal nesta etapa;
- proximas atividades e proximas apresentacoes nao devem virar abas editoriais principais aqui;
- no futuro, uma rotina deve montar snapshot publico, como `publicAgenda/upcoming`, a partir de fontes oficiais ja sanitizadas, sem RGA, e-mail privado, presencas, faltas, justificativas ou observacoes internas.

Leitura publica sanitizada:

- resolve abas pelo Registry usando as keys `PORTAL_PUBLIC_*`;
- le por cabecalho, nunca por indice fixo;
- considera publicaveis somente linhas com `ATIVO = SIM`, `PUBLICAR = SIM` e `STATUS_PUBLICACAO = PUBLICADO/PUBLICADA`, quando essas colunas existirem;
- ordena por `ORDEM` ou `ORDEM_PUBLICA`, quando existir;
- retorna arrays vazios para abas vazias;
- nao retorna colunas extras automaticamente;
- sanitiza textos e URLs;
- nao escreve em planilhas;
- usa cache curto para snapshot/linhas no caminho real;
- `PUBLIC_LOG_PUBLICACAO` nao e exposto pela leitura publica sanitizada.

Contrato do snapshot editorial:

```javascript
{
  ok: true,
  data: {
    pages: {
      home: { blocos: [], atualizadoEm: "" },
      sobre: { blocos: [], atualizadoEm: "" },
      historia: { marcos: [], atualizadoEm: "" },
      parceiros: { itens: [], atualizadoEm: "" }
    },
    documents: [],
    media: [],
    config: {},
    boardComplements: [],
    peopleComplements: [],
    managementComplements: [],
    peopleConfig: {}
  },
  meta: {
    origem: "GEAPA_CORE",
    fonte: "PORTAL_CONTEUDO_PUBLICO",
    atualizadoEm: ""
  }
}
```

Diretoria:

- `corePortalPublicContentGetPessoasComplementos()` retorna complementos editoriais publicos de `PUBLIC_PESSOAS_COMPLEMENTOS`;
- `corePortalPublicContentGetGestoesComplementos()` retorna complementos editoriais publicos de `PUBLIC_GESTOES_COMPLEMENTOS`;
- `corePortalPublicContentGetPessoasConfig()` retorna configuracoes publicas de `PUBLIC_PESSOAS_CONFIG`;
- `corePortalPublicContentGetDiretoriaComplementos()` existe apenas como fallback legado para `PUBLIC_DIRETORIA_COMPLEMENTOS`;
- nao monta a diretoria completa nesta etapa, porque isso exigiria cruzamento com Vigencias/Pessoas;
- a funcao futura documentada para isso e `corePortalPublicContentBuildPublicBoard(options)`.

Observacao de seguranca:

- o navegador nunca deve chamar essa funcao diretamente;
- quem chama e o backend Apps Script do `geapa-portal`;
- o codigo de acesso deve ser enviado sempre para `emailCadastrado` retornado pelo core, nunca para o e-mail digitado pelo usuario quando houver divergencia.

Contrato da tela "Minha situacao":

```javascript
{
  ok: true,
  membro: {
    id: string,
    nomeExibicao: string,
    emailCadastrado: string,
    rga: string,
    vinculo: string,
    situacaoGeral: string
  },
  minhaSituacao: {
    resumo: {
      frequencia: string,
      pendenciasAbertas: number,
      certificadosDisponiveis: number
    },
    pendencias: [
      {
        tipo: "cadastro" | "administrativo",
        titulo: string,
        descricao: string,
        severidade: "baixa" | "media" | "alta",
        status: "pendente"
      }
    ],
    participacao: {
      frequenciaGeral: string,
      atividadesRecentes: [],
      apresentacoes: {
        periodoUltimaApresentacao: string,
        quantidadeRealizadas: number
      }
    },
    diretoria: {
      statusElegibilidade: string,
      diasComputados: number,
      limiteDias: number,
      saldoDias: number,
      dataLimiteEstimada: string
    },
    certificados: [],
    avisos: []
  },
  usuario: { ... } // mesmo contrato de geapaCoreBuscarUsuarioPortal(emailOuRga), quando disponivel
}
```

Retornos controlados:

- membro nao encontrado: `{ ok: false, code: "MEMBRO_NAO_ENCONTRADO", message: "Membro nao encontrado para o e-mail ou RGA informado." }`
- erro inesperado em usuario: `{ ok: false, code: "ERRO_BUSCAR_USUARIO_PORTAL", message: "Nao foi possivel buscar o usuario do portal." }`
- erro inesperado: `{ ok: false, code: "ERRO_BUSCAR_MINHA_SITUACAO", message: "Nao foi possivel buscar a situacao do membro." }`

Pendencias retornadas nesta etapa:

- e-mail cadastrado ausente ou invalido;
- RGA nao informado;
- nome de exibicao nao informado;
- vinculo cadastral indefinido;
- situacao geral indefinida.

Regras das pendencias:

- `resumo.pendenciasAbertas` sempre acompanha o tamanho de `minhaSituacao.pendencias`;
- as mensagens sao amigaveis e nao incluem valores brutos ausentes ou invalidos;
- a funcao continua retornando apenas dados do proprio membro localizado.

Bloco de participacao por apresentacoes:

- `minhaSituacao.participacao.apresentacoes.periodoUltimaApresentacao` vem de `PERIODO_ULTIMA_APRESENTACAO`;
- `minhaSituacao.participacao.apresentacoes.quantidadeRealizadas` vem de `QTD_APRESENTACOES_REALIZADAS`;
- `QTD_APRESENTACOES_REALIZADAS` ja e o total consolidado entre a base legado e as apresentacoes atuais;
- os campos `*_BASE_LEGADO` nao sao expostos ao portal para evitar dupla contagem ou interpretacao ambigua;
- periodos vazios retornam string vazia;
- quantidades vazias, invalidas ou nao numericas retornam `0`.

Bloco orientativo de elegibilidade para Diretoria:

- `minhaSituacao.diretoria.statusElegibilidade` vem de `STATUS_ELEGIBILIDADE_DIRETORIA`;
- `minhaSituacao.diretoria.diasComputados` vem de `QTD_DIAS_QUE_CONTAM_PARA_LIMITE_DIRETORIA`;
- `minhaSituacao.diretoria.limiteDias` vem de `LIMITE_DIAS_DIRETORIA`;
- `minhaSituacao.diretoria.saldoDias` vem de `SALDO_DIAS_DIRETORIA`;
- `minhaSituacao.diretoria.dataLimiteEstimada` vem de `DATA_LIMITE_ESTIMADA_DIRETORIA`;
- status vazio retorna string vazia, sem inventar valor;
- numeros vazios, invalidos, nao numericos ou negativos retornam `0`;
- a data e retornada como texto exibido na planilha, sem conversao para `Date`;
- essa informacao e orientativa; decisoes finais continuam sendo da Diretoria.

Fora de escopo nesta etapa:

- pendencias disciplinares;
- observacoes internas;
- motivos de suspensao ou desligamento;
- avaliacoes subjetivas;
- documentos obrigatorios sem fonte oficial objetiva e nao sensivel no Core.
- frequencia detalhada, lista de presenca e observacoes internas.
- historico de cargos, justificativas internas e detalhes sensiveis de elegibilidade.

Teste manual pelo editor do Apps Script:

1. configure a Script Property `GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR` com um e-mail ou RGA de teste;
2. execute `geapaCoreRunTesteUsuarioPortal()`, `geapaCoreRunTesteListarMembrosParaChamada()` ou `geapaCoreRunTesteMinhaSituacaoParaPortal()`;
3. confira o retorno no log/execucao sem adicionar e-mail real fixo ao codigo.

Campos ainda vazios nesta V1:

- `minhaSituacao.resumo.frequencia`;
- `minhaSituacao.participacao.frequenciaGeral`;
- `minhaSituacao.participacao.atividadesRecentes`;
- `minhaSituacao.certificados`;
- `minhaSituacao.avisos`.

Esses blocos permanecem vazios ou zerados ate haver fonte oficial confiavel integrada ao Core para frequencia, pendencias, certificados e atividades recentes.

### Logs e utilidades

- `coreRunId()`
- `coreLogInfo(runId, message, meta)`
- `coreLogWarn(runId, message, meta)`
- `coreLogError(runId, message, meta)`
- `coreLogSummarize(message, meta)`
- `coreAssertRequired(value, label)`
- utilitarios `Drive` e `HTTP` usados por modulos consumidores.

---

## Keys institucionais mais usadas

O core depende do Registry e, conforme a funcao chamada, pode acessar chaves como:

- `MEMBERS_ATUAIS`
- `CARGOS_INSTITUCIONAIS_CONFIG`
- `VIGENCIA_DIRETORIAS`
- `VIGENCIA_MEMBROS_DIRETORIAS`
- `VIGENCIA_ASSESSORES`
- `VIGENCIA_CONSELHEIROS`
- `VIGENCIA_SEMESTRES`
- `MAIL_EVENTOS`
- `MAIL_INDICE`
- `MAIL_SAIDA`
- `MAIL_ANEXOS`
- `MAIL_REGRAS`
- `MAIL_CONFIG`
- `CONFIG_GEAPA` ou chave legada `DADOS_OFICIAIS_GEAPA`

Modulos consumidores podem acessar outras `KEYS` via `coreGetSheetByKey`, desde que estejam cadastradas no Registry.

---

## Trigger do core

Arquivo de trigger:

- `90_core_triggers.gs`

Funcoes principais:

- `core_installTriggers()`
- `core_reinstallTriggers()`
- `core_uninstallTriggers()`
- `core_listTriggers()`
- `core_validateTriggers()`
- `coreSyncMembersCurrentDerivedFields()`
- `coreMailProcessOutbox()`
- `coreMailIngestInbox(opts)`
- `coreMailCleanupNoiseEvents()`

Uso atual:

- trigger temporal diario para sincronizar campos derivados em `MEMBERS_ATUAIS`.
- entre os derivados sincronizados, o core atualiza o semestre atual, o numero de semestres no grupo e, quando a coluna existir, `TEMPO_EFETIVO_NO_GRUPO` com base em `Data integração`.
- trigger horario para processar a `MAIL_SAIDA`.
- trigger horario para ingestao automatica da caixa de entrada do Mail Hub.
- trigger diario para limpeza de eventos ignorados/ruido no Mail Hub.

---

## Observacoes de manutencao

- mudancas no Registry impactam todos os modulos consumidores;
- as APIs publicas devem ser adicionadas em `20_public_exports.js`; implementar a funcao interna sem exporta-la nao basta para uso via Library;
- modulos que usam Library em versao fixa precisam atualizar a versao publicada apos mudancas no core;
- sempre que um helper novo for usado por outro modulo, documentar o contrato publico correspondente neste README.

Fluxo manual sugerido para validar o Mail Hub:

1. confira se as abas `MAIL_EVENTOS`, `MAIL_INDICE`, `MAIL_ANEXOS` e `MAIL_CONFIG` possuem os cabecalhos minimos acima;
2. rode `test_core_mailHub_assertSchema()` para validar a estrutura;
3. rode `test_core_mailHub_ingestInbox_dryRun()` para confirmar query e volume sem gravar nada;
4. rode `test_core_mailHub_ingestInbox_real()` para gravar eventos reais;
5. confira as abas `Eventos de Email`, `Indice de Conversas` e `Anexos`;
6. rode `test_core_mailHub_listPending_membros()` ou `test_core_mailHub_listPending_naoIdentificado()` para validar a consulta de pendencias.
7. use `test_core_mailHub_getLatestEvent()` ou `test_core_mailHub_getLatestPending_membros()` para localizar rapidamente o ultimo evento de teste.
8. use `test_core_mailHub_markLatestPending_membros_processed()` para marcar o ultimo pendente sem copiar `eventId` na mao.
9. para validar a fila central, rode `test_core_mailOutbox_queue_operacional()` e confira a nova linha em `MAIL_SAIDA`.
10. em seguida rode `test_core_mailOutbox_process()` e confirme `Status Envio = ENVIADO`, `Enviado Em`, `Id Thread Gmail`, `Id Mensagem Gmail` e o reflexo em `MAIL_EVENTOS` / `MAIL_INDICE`.

### Cache Firestore do login do Portal

O GEAPA-CORE pode gerar snapshots seguros para `portalUsers/{uid}` usando PESSOAS v2 como fonte oficial. O Firestore e apenas cache operacional do Portal: o documento deve conter somente os campos minimos de interface (`uid`, `idPessoa`, `nomeExibicao`, `email`, `rga`, `portalAtivo`, perfis, permissoes, vinculo atual, `source`, `sourceUpdatedAt`, `cacheUpdatedAt`, `cacheExpiresAt` e `schemaVersion`).

Funcoes principais: `corePortalGerarSnapshotFirestoreUsuario`, `corePortalSincronizarUsuarioFirestore`, `corePortalInvalidarCacheFirestoreUsuario`, `corePortalSyncFirestoreUserByEmail`, `corePortalSyncFirestoreUserByIdPessoa` e `corePortalSyncFirestoreUsersFromPessoasV2`. A escrita atual usa Apps Script + Firestore REST no plano Spark, configurada por Script Properties (`GEAPA_CORE_FIRESTORE_PROJECT_ID` e opcionalmente `GEAPA_CORE_FIRESTORE_DATABASE_ID`), sem Cloud Functions, Secret Manager, service account ou segredo no repositorio.

O transporte REST compartilhado fica em `24_core_firestore_rest.js` e tambem
expoe `coreFirestoreSetDocument`, `coreFirestoreGetDocument`,
`coreFirestoreListDocuments`, `coreFirestoreDeleteDocument`,
`coreFirestoreBatchSetDocuments` e `coreFirestoreDiagnosticar`. Essas funcoes
usam `dryRun: true` por padrao e suportam os read models de Atividades sem
transformar o Firestore em fonte oficial. Consulte
`docs/firestore-read-models-atividades.md`.
