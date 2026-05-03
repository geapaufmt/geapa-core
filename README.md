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
- o slogan exibido no rodape nao e fixo: ele e buscado da coluna `Slogan` da diretoria vigente em `VIGENCIA_DIRETORIAS`, com fallback seguro quando estiver vazio.
- a identidade oficial do grupo usada no renderer passa a ser lida de `DADOS_OFICIAIS_GEAPA`, incluindo nome oficial, sigla, e-mail oficial e cores institucionais;
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
- quando o modulo optar por envio em massa via `bcc`, o core usa `EMAIL_OFICIAL` em `DADOS_OFICIAIS_GEAPA` como envelope principal de seguranca.

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
- `geapaCoreBuscarMinhaSituacaoParaPortal(emailOuRga)`
- `geapaCoreRunTesteMinhaSituacaoParaPortal()`

Regras do contrato:

- aceita e-mail ou RGA;
- normaliza entrada com `trim` e, para e-mail, `lowercase`;
- consulta `MEMBERS_ATUAIS` via Registry;
- a consulta cadastral retorna um unico membro ou `null`;
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
        quantidadeRealizadas: number,
        periodoUltimaApresentacaoBaseLegado: string,
        quantidadeRealizadasBaseLegado: number
      }
    },
    certificados: [],
    avisos: []
  }
}
```

Retornos controlados:

- membro nao encontrado: `{ ok: false, code: "MEMBRO_NAO_ENCONTRADO", message: "Membro nao encontrado para o e-mail ou RGA informado." }`
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
- `minhaSituacao.participacao.apresentacoes.periodoUltimaApresentacaoBaseLegado` vem de `PERIODO_ULTIMA_APRESENTACAO_BASE_LEGADO`;
- `minhaSituacao.participacao.apresentacoes.quantidadeRealizadasBaseLegado` vem de `QTD_APRESENTACOES_REALIZADAS_BASE_LEGADO`;
- periodos vazios retornam string vazia;
- quantidades vazias, invalidas ou nao numericas retornam `0`.

Fora de escopo nesta etapa:

- pendencias disciplinares;
- observacoes internas;
- motivos de suspensao ou desligamento;
- avaliacoes subjetivas;
- documentos obrigatorios sem fonte oficial objetiva e nao sensivel no Core.
- frequencia detalhada, lista de presenca e observacoes internas.

Teste manual pelo editor do Apps Script:

1. configure a Script Property `GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR` com um e-mail ou RGA de teste;
2. execute `geapaCoreRunTesteMinhaSituacaoParaPortal()`;
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
- `DADOS_OFICIAIS_GEAPA`

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
