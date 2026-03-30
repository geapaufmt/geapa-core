# GEAPA Core (Apps Script Library)

Biblioteca compartilhada do ecossistema GEAPA. O `geapa-core` centraliza acesso a planilhas por Registry, normalizacao de dados, servicos de e-mail/Gmail, leitura tabular orientada a cabecalhos, governanca institucional e utilitarios reutilizados pelos demais modulos.

---

## Objetivo

O core existe para:

- resolver planilhas por `KEY` institucional via Registry;
- expor uma API publica estavel para os modulos consumidores;
- centralizar leitura/escrita por cabecalho e por registros;
- unificar envio de e-mails, replies e rastreamento de threads;
- projetar ocupantes atuais de cargos/funcoes institucionais;
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

### Sheets e records

Camada reutilizavel para leitura e escrita sem depender de colunas fixas.

Funcoes centrais:

- `coreNormalizeHeader(value)`
- `coreBuildHeaderIndexMap(headers, opts)`
- `coreFindHeaderIndex(headerMap, headerName, opts)`
- `coreGetCellByHeader(row, headerMap, headerName, opts)`
- `coreSetRowValueByHeader(row, headerMap, headerName, value, opts)`
- `coreWriteCellByHeader(sheet, rowNumber, headerMap, headerName, value, opts)`
- `coreAppendObjectByHeaders(sheet, payload, opts)`
- `coreReadSheetRecords(sheet, opts)`
- `coreReadRecordsByKey(key, opts)`
- `coreFindFirstRecordByField(records, headerName, value, opts)`
- `coreFindFirstRecordByAnyField(records, headerNames, value, opts)`
- `coreGetNearestFilledValueUp(sheet, rowNumber, colNumber)`

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

### Mail Hub (V1)

Camada central de ingestao de e-mails do Gmail para registro institucional em planilhas centrais.

Escopo atual:

- ler mensagens do Gmail com deduplicacao por `Id Mensagem Gmail`;
- extrair `Chave de Correlacao` do assunto no padrao `[GEAPA][CHAVE]`;
- resolver `Modulo Dono` e metadados de roteamento via adapters por modulo;
- registrar eventos em `MAIL_EVENTOS`;
- manter upsert de indice em `MAIL_INDICE`;
- registrar metadados de anexos em `MAIL_ANEXOS`;
- consultar pendencias por modulo e marcar evento como processado.

Funcoes publicas:

- `coreMailRegisterModuleAdapter(adapter)`
- `coreMailGetModuleAdapter(moduleCodeOrName)`
- `coreMailListModuleAdapters()`
- `coreMailBuildCorrelationKey(moduleCodeOrName, ctx)`
- `coreMailParseCorrelationKey(key)`
- `coreMailResolveRouting(msgCtx)`
- `coreMailNormalizeOutgoingSubject(moduleCodeOrName, subject, ctx)`
- `coreMailIngestInbox(opts)`
- `coreMailGetConfig(key, defaultValue)`
- `coreMailGetConfigBoolean(key, defaultValue)`
- `coreMailGetConfigList(key)`
- `coreMailListPendingByModule(moduleName)`
- `coreMailGetLatestEvent(opts)`
- `coreMailMarkLatestPendingByModule(moduleName, processorName)`
- `coreMailMarkEventProcessed(eventId, processorName)`
- `coreMailCleanupNoiseEvents()`

Arquitetura de adapters:

- o Mail Hub nao conhece regra de negocio de `APRESENTACOES`, `SELETIVO`, `MEMBROS` ou outros modulos;
- cada adapter declara apenas um contrato minimo comum: construcao de chave, parsing, match, resolucao de roteamento e normalizacao opcional de assunto;
- o core faz registry dos adapters e escolhe o melhor roteamento com base em `correlationKey` ou heuristicas do proprio adapter;
- adapters iniciais incluidos no core: `APR / APRESENTACOES`, `SEL / SELETIVO`, `MEM / MEMBROS`;
- o contrato permite formatos diferentes de chave entre modulos, desde que o proprio adapter saiba construir e interpretar a sua chave.

Observacao de plataforma:

- em Apps Script, registrar adapters com funcoes a partir de projetos consumidores via Library pode ter limitacoes de callbacks; a implementacao atual e segura para adapters definidos no proprio core ou no mesmo projeto.

Observacoes desta V1:

- nao migra os modulos consumidores existentes;
- nao implementa fila de saida;
- nao salva anexos no Drive;
- nao aplica roteamento avancado por regras.

Higiene de ingestao e consistencia semantica:

- o hub le `MAIL_CONFIG` por chave e aplica regras de ingestao sem hardcode operacional;
- remetentes, dominios e assuntos tecnicos podem ser ignorados antes do registro;
- se `USAR_SOMENTE_ASSUNTOS_GEAPA = SIM`, o assunto precisa respeitar `ASSUNTO_PREFIXO_OBRIGATORIO`;
- ruido tecnico de GitHub, Codex e alertas similares e preferencialmente descartado;
- se `MARCAR_RUIDO_COMO_IGNORADO = SIM`, o ruido passa a ser registrado com `Status Roteamento = IGNORADO` e `Status Processamento = IGNORADO`;
- o indice e recomposto por resumo da chave de correlacao, preenchendo entidade, etapa, contadores e flags pendentes.

Schema minimo esperado na versao atual da planilha central:

- `MAIL_EVENTOS`: `Id Evento`, `Data Hora Evento`, `Direcao`, `Tipo Evento`, `Modulo Dono`, `Chave de Correlacao`, `Id Thread Gmail`, `Id Mensagem Gmail`, `Assunto`, `Email Remetente`, `Emails Destinatarios`, `Status Processamento`, `Processado Por`, `Data Hora Processamento`, `Possui Anexos`, `Quantidade Anexos`, `Criado Em`, `Atualizado Em`
- `MAIL_INDICE`: `Chave de Correlacao`, `Modulo Dono`, `Tipo Entidade`, `Id Entidade`, `Etapa Atual`, `Id Thread Gmail`, `Id Ultima Mensagem`, `Ultima Direcao`, `Ultimo Tipo Evento`, `Ultimo Email Remetente`, `Ultimo Assunto`, `Data Hora Ultimo Evento`, `Ha Entrada Pendente`, `Ha Anexo Pendente`, `Quantidade Eventos`, `Quantidade Entradas`, `Quantidade Saidas`, `Quantidade Anexos`, `Criado Em`, `Atualizado Em`
- `MAIL_ANEXOS`: `Id Anexo`, `Id Evento`, `Modulo Dono`, `Chave de Correlacao`, `Etapa Fluxo`, `Id Mensagem Gmail`, `Id Thread Gmail`, `Nome Arquivo`, `Tipo Mime`, `Tamanho Bytes`, `Status Anexo`, `Criado Em`, `Atualizado Em`
- `MAIL_CONFIG`: `Chave`, `Valor`, `Ativo`

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

Testes manuais no projeto:

- `test_core_mailAdapters_list()`
- `test_core_mailAdapters_get_mem()`
- `test_core_mailAdapters_build_mem()`
- `test_core_mailAdapters_parse_mem()`
- `test_core_mailAdapters_resolveRouting_mem()`
- `test_core_mailAdapters_normalizeOutgoingSubject_mem()`
- `test_core_mailHub_assertSchema()`
- `test_core_mailHub_config_read()`
- `test_core_mailHub_ingestInbox_dryRun()`
- `test_core_mailHub_ingestInbox_higiene_dryRun()`
- `test_core_mailHub_ingestInbox_real()`
- `test_core_mailHub_listPending_naoIdentificado()`
- `test_core_mailHub_listPending_membros()`
- `test_core_mailHub_getLatestEvent()`
- `test_core_mailHub_getLatestPending_membros()`
- `test_core_mailHub_markLatestPending_membros_processed()`
- `test_core_mailHub_markLatestPending_naoIdentificado_processed()`
- `test_core_mailHub_cleanupNoiseEvents()`

Sugestoes de melhoria na planilha:

- adicionar validacao de dados em `Direcao`, `Tipo Evento`, `Status Roteamento`, `Status Processamento`, `Status Conversa` e `Status Anexo` para evitar variacoes de texto;
- congelar a linha 1 em todas as abas e manter os nomes dos cabecalhos como estao hoje, porque o codigo depende deles por nome;
- na aba `Configuracoes`, preencher `GMAIL_QUERY_INGEST`, `GMAIL_MAX_THREADS` e `GMAIL_MAX_MESSAGES_PER_THREAD` para controlar a ingestao sem mexer no codigo;
- se quiser facilitar operacao humana, aplicar filtros nas abas `Eventos de Email`, `Indice de Conversas` e `Anexos`;
- se quiser ganhar rastreabilidade futura, vale considerar uma coluna opcional `Correlation Prefix` ou `Prefixo Correlacao` em `Indice de Conversas`, mas ela nao e necessaria para a V1 funcionar.

### Governanca institucional

Projecao de ocupantes atuais por cargo/funcao e por grupo de e-mails.

Funcoes centrais:

- `coreGetCurrentBoard(refDate)`
- `coreGetCurrentBoardMembers(refDate)`
- `coreGetCurrentBoardMemberByRole(role, refDate)`
- `coreGetCurrentLeadership(refDate)`
- `coreGetInstitutionalRolesActive()`
- `coreFindInstitutionalRoleByAnyName(text)`
- `coreGetCurrentInstitutionalAssignments(refDate)`
- `coreGetCurrentOccupantsByEmailGroup(groupName, refDate)`
- `coreGetCurrentContactsHtmlByEmailGroup(groupName, refDate)`
- `coreGetCurrentEmailsByEmailGroup(groupName, refDate)`
- `coreGetCurrentEmailsByRole(roleName, refDate)`

### Identidade de membros

Busca e autofill com base em `MEMBERS_ATUAIS`.

Funcoes centrais:

- `coreNormalizeIdentityKey(value)`
- `coreFindMemberIdentityByAny(identity)`
- `coreFindMemberCurrentRowByAny(identity)`
- `coreAutofillIdentityRowInSheet(sheet, rowNumber, opts)`

Observacao:

- `coreAutofillIdentityRowInSheet` aceita `opts.nameHeaders`, `opts.rgaHeaders` e `opts.emailHeaders` para modulos com cabecalhos especificos, preservando o comportamento padrao quando `opts` nao e informado.

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

Modulos consumidores podem acessar outras `KEYS` via `coreGetSheetByKey`, desde que estejam cadastradas no Registry.

---

## Trigger do core

Arquivo de trigger:

- `90_core_triggers.gs`

Funcoes principais:

- `core_installTriggers()`
- `core_uninstallTriggers()`
- `coreSyncMembersCurrentDerivedFields()`

Uso atual:

- trigger temporal diario para sincronizar campos derivados em `MEMBERS_ATUAIS`.

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
