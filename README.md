# GEAPA Core (Apps Script Library)

Biblioteca base do ecossistema GEAPA para padronizar acesso a dados institucionais,
projeção de cargos/funções e sincronizações operacionais em planilhas.

---

## 1) Objetivo

Este módulo existe para:

- resolver planilhas por `KEY` institucional via Registry;
- centralizar regras de governança/vigências;
- projetar cargos atuais (diretoria, assessoria, conselho);
- sincronizar campos derivados em `MEMBERS_ATUAIS`;
- oferecer utilitários reutilizáveis (email, drive, logs, datas, assert, http).

---

## 2) Fluxo principal

### 2.1 Registry → referência de planilha
1. Lê planilha-mãe do Registry (aba `Registry`).
2. Valida cabeçalhos obrigatórios.
3. Interpreta `ATIVO` (`SIM` / `NÃO`) e `AMBIENTE` (`DEV` / `PROD`).
4. Retorna mapa efetivo por ambiente: `KEY -> { id, sheet }`.

### 2.2 Projeção institucional
1. Lê configuração em `CARGOS_INSTITUCIONAIS_CONFIG`.
2. Lê vigências (`VIGENCIA_MEMBROS_DIRETORIAS`, `VIGENCIA_ASSESSORES`, `VIGENCIA_CONSELHEIROS`).
3. Normaliza textos e datas.
4. Gera atribuições atuais e agrupamentos por pessoa.

### 2.3 Sincronização de MEMBERS_ATUAIS
1. Atualiza `Cargo/função atual`.
2. Atualiza derivados acadêmicos (`Semestre atual`, `N° de semestres no grupo`).
3. Retorna resumo de atualização (linhas/células alteradas).

---

## 3) Gatilhos

Arquivo: `90_core_triggers.gs`

- `core_installTriggers()`
  - remove trigger antigo de `coreSyncMembersCurrentDerivedFields`;
  - cria trigger diário às 06:00.
- `core_uninstallTriggers()`
  - remove os triggers desse handler.

Handler executado por trigger:
- `coreSyncMembersCurrentDerivedFields()`.

---

## 4) Planilhas esperadas (KEYS do Registry)

- `MEMBERS_ATUAIS`
- `CARGOS_INSTITUCIONAIS_CONFIG`
- `VIGENCIA_DIRETORIAS`
- `VIGENCIA_MEMBROS_DIRETORIAS`
- `VIGENCIA_ASSESSORES`
- `VIGENCIA_CONSELHEIROS`
- `VIGENCIA_SEMESTRES`

Planilha-mãe:
- Spreadsheet fixo definido no código do registry;
- aba fixa `Registry`.

---

## 5) Colunas/cabeçalhos usados

### 5.1 Aba `Registry`

**Obrigatórios**
- `KEY`
- `SPREADSHEET_ID`
- `SHEET_NAME`
- `ATIVO` (`SIM` / `NÃO` ou `NAO`)
- `AMBIENTE` (`DEV` / `PROD`)

**Opcionais**
- `DISPLAY_NAME`
- `TYPE`
- `NOTAS`

### 5.2 `CARGOS_INSTITUCIONAIS_CONFIG`

Cabeçalhos obrigatórios:
- `NOME_PUBLICO`
- `GRUPO_CARGO`
- `CARGO_KEY`
- `ATIVO`
- `DISPLAY_ORDEM`
- `RECEBE_EMAILS`
- `EMAILS_GRUPO`
- `É_CARGO_UNICO`
- `ESCRITA_VARIACAO`

### 5.3 `MEMBERS_ATUAIS`

Usados por diferentes rotinas:
- `MEMBRO` (ou `Nome`)
- `RGA`
- `EMAIL` (ou `E-mail`)
- `TELEFONE` (ou `Telefone`)
- `Status`
- `Cargo/função atual`
- `Semestre de Entrada`
- `Semestre atual`
- `N° de semestres no grupo`

### 5.4 Vigências

**`VIGENCIA_DIRETORIAS`**
- `ID_Diretoria`
- `Início_Mandato`
- `Fim_Mandato`

**`VIGENCIA_MEMBROS_DIRETORIAS`**
- `Nome`
- `RGA`
- `Cargo/Função`
- `ID_Diretoria`
- `Data_Início`
- `Data_Fim`
- `Data_Fim_previsto`

**`VIGENCIA_ASSESSORES` / `VIGENCIA_CONSELHEIROS`**
- parser aceita aliases para nome, email, cargo, datas e ativo.

### 5.5 `VIGENCIA_SEMESTRES`

- `ID_Semestre` (esperado `YYYY/S`)
- `Início`
- `Fim`
- `ID_Período` (opcional conforme fluxo)

---

## 6) Funções principais (API pública)

### 6.1 Registry e planilhas
- `coreGetRegistry()`
- `coreGetRegistryRefByKey(key)`
- `coreGetSheetByKey(key)`
- `coreGetRegistryMetaByKey(key)`
- `coreGetCurrentEnv()`
- `coreClearRegistryCache()`

### 6.2 Governança/cargos/projeção
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

### 6.3 Sincronização e semestre
- `coreSyncMembersCurrentInstitutionalRoles(refDate)`
- `coreSyncMembersCurrentDerivedFields()`
- `coreGetCurrentSemester(refDate)`
- `coreParseEntrySemesterFromRga(rga)`
- `coreGetStudentCurrentSemesterFromRga(rga, refDate)`
- `coreGetSemesterForDate(refDate)`
- `coreGetLastCompletedSemester(refDate)`
- `coreGetCompletedGroupSemesterCountFromEntrySemester(entrySemesterShort, refDate)`

### 6.4 Identidade de membro
- `coreFindMemberIdentityByAny(identity)`
- `coreAutofillIdentityRowInSheet(sheet, rowNumber)`

### 6.5 Serviços de apoio
- `coreSendEmailText`, `coreSendEmailHtml`, `coreSendHtmlEmail`
- `coreDrive*`
- `coreHttpPostJson`
- `coreRunId`, `coreLogInfo`, `coreLogWarn`, `coreLogError`, `coreLogSummarize`
- `coreAssertRequired`

---

## 7) Pontos de manutenção futura

1. **Validação formal de schema por KEY**
   - criar checks centralizados por planilha para detectar quebra de cabeçalho cedo.

2. **Padronização de cabeçalhos**
   - reduzir aliases (`MEMBRO`/`Nome`, `EMAIL`/`E-mail`) para diminuir complexidade.

3. **Testes automatizados**
   - priorizar testes para parsers de data/semestre, projeções e sincronizações.

4. **Observabilidade operacional**
   - ampliar logs de sync por tipo de atualização e por fonte.

5. **Governança de mudanças no Registry**
   - mudanças em `ATIVO`/`AMBIENTE` impactam execução global; usar checklist de alteração.
