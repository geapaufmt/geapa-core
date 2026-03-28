# GEAPA Core (Apps Script Library)

Biblioteca compartilhada do ecossistema GEAPA. O `geapa-core` centraliza acesso a planilhas por Registry, normalização de dados, serviços de e-mail/Gmail, leitura tabular orientada a cabeçalhos, governança institucional e utilitários reutilizados pelos demais módulos.

---

## Objetivo

O core existe para:

- resolver planilhas por `KEY` institucional via Registry;
- expor uma API pública estável para os módulos consumidores;
- centralizar leitura/escrita por cabeçalho e por registros;
- unificar envio de e-mails, replies e rastreamento de threads;
- projetar ocupantes atuais de cargos/funções institucionais;
- sincronizar campos derivados em `MEMBERS_ATUAIS`;
- oferecer utilitários comuns de texto, datas, identidade e logs.

---

## Áreas principais

### Registry e acesso a planilhas

- resolve `KEY -> { spreadsheetId, sheetName }` conforme ambiente;
- abre `Sheet` diretamente via `coreGetSheetByKey(key)`;
- mantém cache de registry por execução.

Funções centrais:

- `coreGetRegistry()`
- `coreGetRegistryRefByKey(key)`
- `coreGetSheetByKey(key)`
- `coreGetRegistryMetaByKey(key)`
- `coreClearRegistryCache()`

### Sheets e records

Camada reutilizável para leitura e escrita sem depender de colunas fixas.

Funções centrais:

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

Funções para datas operacionais e leitura da vigência de semestres.

Funções centrais:

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

Funções centrais:

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

### Governança institucional

Projeção de ocupantes atuais por cargo/função e por grupo de e-mails.

Funções centrais:

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

Funções centrais:

- `coreNormalizeIdentityKey(value)`
- `coreFindMemberIdentityByAny(identity)`
- `coreFindMemberCurrentRowByAny(identity)`
- `coreAutofillIdentityRowInSheet(sheet, rowNumber, opts)`

Observação:

- `coreAutofillIdentityRowInSheet` aceita `opts.nameHeaders`, `opts.rgaHeaders` e `opts.emailHeaders` para módulos com cabeçalhos específicos, preservando o comportamento padrão quando `opts` não é informado.

### Logs e utilidades

- `coreRunId()`
- `coreLogInfo(runId, message, meta)`
- `coreLogWarn(runId, message, meta)`
- `coreLogError(runId, message, meta)`
- `coreLogSummarize(message, meta)`
- `coreAssertRequired(value, label)`
- utilitários `Drive` e `HTTP` usados por módulos consumidores.

---

## Keys institucionais mais usadas

O core depende do Registry e, conforme a função chamada, pode acessar chaves como:

- `MEMBERS_ATUAIS`
- `CARGOS_INSTITUCIONAIS_CONFIG`
- `VIGENCIA_DIRETORIAS`
- `VIGENCIA_MEMBROS_DIRETORIAS`
- `VIGENCIA_ASSESSORES`
- `VIGENCIA_CONSELHEIROS`
- `VIGENCIA_SEMESTRES`

Módulos consumidores podem acessar outras `KEYS` via `coreGetSheetByKey`, desde que estejam cadastradas no Registry.

---

## Trigger do core

Arquivo de trigger:

- `90_core_triggers.gs`

Funções principais:

- `core_installTriggers()`
- `core_uninstallTriggers()`
- `coreSyncMembersCurrentDerivedFields()`

Uso atual:

- trigger temporal diário para sincronizar campos derivados em `MEMBERS_ATUAIS`.

---

## Observações de manutenção

- Mudanças no Registry impactam todos os módulos consumidores.
- As APIs públicas devem ser adicionadas em `20_public_exports.js`; implementar a função interna sem exportá-la não basta para uso via Library.
- Módulos que usam Library em versão fixa precisam atualizar a versão publicada após mudanças no core.
- Sempre que um helper novo for usado por outro módulo, documentar o contrato público correspondente neste README.