/***************************************
 * 13_core_governance.gs
 *
 * Camada institucional de vigências/diretoria do GEAPA.
 *
 * Objetivos:
 * - Descobrir qual diretoria está vigente
 * - Retornar os membros da diretoria vigente
 * - Cruzar dados da diretoria com MEMBERS_ATUAIS
 * - Permitir consulta por cargo/função atual
 *
 * Dependências:
 * - Registry:
 *   - VIGENCIA_DIRETORIAS
 *   - VIGENCIA_MEMBROS_DIRETORIAS
 *   - MEMBERS_ATUAIS
 *
 * Observação:
 * - Os contatos (telefone/email) são buscados em MEMBERS_ATUAIS via RGA.
 ***************************************/

const CORE_GOVERNANCE_CFG = Object.freeze({
  boardKey: "VIGENCIA_DIRETORIAS",
  boardMembersKey: "VIGENCIA_MEMBROS_DIRETORIAS",
  membersKey: "MEMBERS_ATUAIS",

  boardHeaders: Object.freeze({
    boardId: "ID_Diretoria",
    start: "Início_Mandato",
    end: "Fim_Mandato"
  }),

  boardMemberHeaders: Object.freeze({
    name: "Nome",
    rga: "RGA",
    role: "Cargo/Função",
    boardId: "ID_Diretoria",
    start: "Data_Início",
    end: "Data_Fim",
    expectedEnd: "Data_Fim_previsto"
  }),

  membersHeaders: Object.freeze({
    name: "MEMBRO",
    rga: "RGA",
    phone: "TELEFONE",
    email: "EMAIL",
    currentRole: "Cargo/função atual"
  })
});


/* ======================================================================
 * Helpers
 * ====================================================================== */

function core_normalizeGovernanceText_(value) {
  return String(value || "").trim().toLowerCase();
}

function core_parseDateOrNull_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) return value;

  const d = new Date(value);
  return isNaN(d) ? null : d;
}

function core_isDateInRange_(targetDate, startDate, endDate) {
  const t = core_parseDateOrNull_(targetDate);
  const start = core_parseDateOrNull_(startDate);
  const end = core_parseDateOrNull_(endDate);

  if (!t || !start) return false;
  if (t < start) return false;
  if (end && t > end) return false;
  return true;
}

function core_getHeaderIndexMap_(headers, wantedMap) {
  const normalized = headers.map(core_normalizeGovernanceText_);
  const out = {};

  Object.keys(wantedMap).forEach(key => {
    out[key] = normalized.indexOf(core_normalizeGovernanceText_(wantedMap[key]));
  });

  return out;
}

function core_getRowsFromSheetByKey_(key) {
  const sh = core_getSheetByKey_(key);
  if (!sh) return { headers: [], rows: [] };

  const values = sh.getDataRange().getValues();
  if (!values.length) return { headers: [], rows: [] };

  const headers = values.shift().map(v => String(v || "").trim());
  return { headers, rows: values };
}

function core_buildMembersCurrentMapByRga_() {
  const data = core_getRowsFromSheetByKey_(CORE_GOVERNANCE_CFG.membersKey);
  const idx = core_getHeaderIndexMap_(data.headers, CORE_GOVERNANCE_CFG.membersHeaders);

  if (idx.rga < 0) {
    throw new Error('core_buildMembersCurrentMapByRga_: cabeçalho "RGA" não encontrado em MEMBERS_ATUAIS.');
  }

  const map = {};

  data.rows.forEach(row => {
    const rga = String(row[idx.rga] || "").trim();
    if (!rga) return;

    map[rga] = Object.freeze({
      name: idx.name >= 0 ? String(row[idx.name] || "").trim() : "",
      rga,
      phone: idx.phone >= 0 ? String(row[idx.phone] || "").trim() : "",
      email: idx.email >= 0 ? String(row[idx.email] || "").trim() : "",
      currentRole: idx.currentRole >= 0 ? String(row[idx.currentRole] || "").trim() : ""
    });
  });

  return map;
}


/* ======================================================================
 * Diretoria vigente
 * ====================================================================== */

/**
 * Retorna a diretoria vigente na data informada.
 *
 * @param {Date=} refDate
 * @return {Object|null}
 */
function core_getCurrentBoard_(refDate) {
  const now = refDate || new Date();
  const data = core_getRowsFromSheetByKey_(CORE_GOVERNANCE_CFG.boardKey);
  const idx = core_getHeaderIndexMap_(data.headers, CORE_GOVERNANCE_CFG.boardHeaders);

  if (idx.boardId < 0 || idx.start < 0 || idx.end < 0) {
    throw new Error("core_getCurrentBoard_: cabeçalhos obrigatórios não encontrados em VIGENCIA_DIRETORIAS.");
  }

  for (let i = 0; i < data.rows.length; i++) {
    const row = data.rows[i];
    const boardId = String(row[idx.boardId] || "").trim();
    const start = row[idx.start];
    const end = row[idx.end];

    if (!boardId) continue;
    if (core_isDateInRange_(now, start, end)) {
      return Object.freeze({
        id: boardId,
        startDate: core_parseDateOrNull_(start),
        endDate: core_parseDateOrNull_(end)
      });
    }
  }

  return null;
}

/**
 * Retorna os membros da diretoria vigente, já cruzados com MEMBERS_ATUAIS.
 *
 * @param {Date=} refDate
 * @return {Object[]}
 */
function core_getCurrentBoardMembers_(refDate) {
  const board = core_getCurrentBoard_(refDate);
  if (!board) return [];

  const currentMembersMap = core_buildMembersCurrentMapByRga_();
  const data = core_getRowsFromSheetByKey_(CORE_GOVERNANCE_CFG.boardMembersKey);
  const idx = core_getHeaderIndexMap_(data.headers, CORE_GOVERNANCE_CFG.boardMemberHeaders);

  if (idx.name < 0 || idx.rga < 0 || idx.role < 0 || idx.boardId < 0) {
    throw new Error("core_getCurrentBoardMembers_: cabeçalhos obrigatórios não encontrados em VIGENCIA_MEMBROS_DIRETORIAS.");
  }

  const result = [];

  data.rows.forEach(row => {
    const boardId = String(row[idx.boardId] || "").trim();
    if (boardId !== board.id) return;

    const role = String(row[idx.role] || "").trim();
    const name = String(row[idx.name] || "").trim();
    const rga = String(row[idx.rga] || "").trim();
    const start = idx.start >= 0 ? core_parseDateOrNull_(row[idx.start]) : null;
    const end = idx.end >= 0 ? core_parseDateOrNull_(row[idx.end]) : null;
    const expectedEnd = idx.expectedEnd >= 0 ? core_parseDateOrNull_(row[idx.expectedEnd]) : null;

    const now = refDate || new Date();

    if (start && now < start) return;
    if (end && now > end) return;

    const memberBase = currentMembersMap[rga] || {};

    result.push(Object.freeze({
      boardId,
      name: memberBase.name || name,
      rga,
      role,
      phone: memberBase.phone || "",
      email: memberBase.email || "",
      currentRoleInMembersSheet: memberBase.currentRole || "",
      startDate: start,
      endDate: end,
      expectedEndDate: expectedEnd
    }));
  });

  return result;
}

/**
 * Retorna os membros da diretoria vigente para um cargo/função.
 *
 * Comparação case-insensitive.
 *
 * @param {string} role
 * @param {Date=} refDate
 * @return {Object[]}
 */
function core_getCurrentBoardMembersByRole_(role, refDate) {
  const wanted = core_normalizeGovernanceText_(role);
  if (!wanted) return [];

  return core_getCurrentBoardMembers_(refDate).filter(member => {
    return core_normalizeGovernanceText_(member.role) === wanted;
  });
}

/**
 * Retorna o primeiro membro da diretoria vigente para um cargo/função.
 *
 * @param {string} role
 * @param {Date=} refDate
 * @return {Object|null}
 */
function core_getCurrentBoardMemberByRole_(role, refDate) {
  const found = core_getCurrentBoardMembersByRole_(role, refDate);
  return found.length ? found[0] : null;
}

/**
 * Retorna a liderança principal da diretoria vigente.
 *
 * @param {Date=} refDate
 * @return {Object}
 */
function core_getCurrentLeadership_(refDate) {
  return Object.freeze({
    presidente: core_getCurrentBoardMemberByRole_("Presidente", refDate),
    vicePresidente: core_getCurrentBoardMemberByRole_("Vice Presidente", refDate),
    secretarioGeral: core_getCurrentBoardMemberByRole_("Secretário(a) Geral", refDate),
    secretarioExecutivo: core_getCurrentBoardMemberByRole_("Secretário(a) Executivo", refDate),
    coordenadorComunicacao: core_getCurrentBoardMemberByRole_("Coordenador(a) de Comunicação", refDate)
  });
}