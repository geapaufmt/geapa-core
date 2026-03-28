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

/* ======================================================================
 * Semestre vigente / semestre por RGA
 * ====================================================================== */

/**
 * Retorna o semestre institucional vigente na data informada.
 *
 * Regra:
 * 1. se a data estiver dentro de um semestre, retorna esse semestre
 * 2. se houver buraco entre semestres, retorna o próximo semestre futuro
 *
 * Exemplo de retorno:
 * {
 *   id: "2026/1",
 *   startDate: Date,
 *   endDate: Date,
 *   periodId: "GEAPA_2026"
 * }
 *
 * Fonte:
 * - VIGENCIA_SEMESTRES
 *
 * @param {Date=} refDate
 * @return {Object|null}
 */
function core_getCurrentSemester_(refDate) {
  const now = refDate || new Date();
  const data = core_getRowsFromSheetByKey_("VIGENCIA_SEMESTRES");

  const idx = core_getHeaderIndexMap_(data.headers, {
    semesterId: "ID_Semestre",
    start: "Início",
    end: "Fim",
    periodId: "ID_Período"
  });

  if (idx.semesterId < 0 || idx.start < 0 || idx.end < 0) {
    throw new Error("core_getCurrentSemester_: cabeçalhos obrigatórios não encontrados em VIGENCIA_SEMESTRES.");
  }

  let nextSemester = null;

  for (let i = 0; i < data.rows.length; i++) {
    const row = data.rows[i];
    const semesterId = String(row[idx.semesterId] || "").trim();
    const start = core_parseDateOrNull_(row[idx.start]);
    const end = core_parseDateOrNull_(row[idx.end]);
    const periodId = idx.periodId >= 0 ? String(row[idx.periodId] || "").trim() : "";

    if (!semesterId || !start) continue;

    // Caso 1: semestre vigente de fato
    if (core_isDateInRange_(now, start, end)) {
      return Object.freeze({
        id: semesterId,
        startDate: start,
        endDate: end,
        periodId: periodId
      });
    }

    // Caso 2: guardar o próximo semestre futuro mais próximo
    if (start > now) {
      if (!nextSemester || start < nextSemester.startDate) {
        nextSemester = Object.freeze({
          id: semesterId,
          startDate: start,
          endDate: end,
          periodId: periodId
        });
      }
    }
  }

  // Se não houver semestre vigente, assume o próximo semestre futuro
  if (nextSemester) return nextSemester;

  return null;
}

/**
 * Extrai ano e semestre de ingresso a partir do RGA.
 *
 * Regra esperada:
 * - 4 primeiros dígitos = ano
 * - 5º dígito = semestre de ingresso (1 ou 2)
 *
 * Exemplo:
 * 202311801028 -> { year: 2023, semester: 1 }
 *
 * @param {string|number} rga
 * @return {Object|null}
 */
function core_parseEntrySemesterFromRga_(rga) {
  const digits = String(rga || "").replace(/\D/g, "");
  if (digits.length < 5) return null;

  const year = parseInt(digits.slice(0, 4), 10);
  const semester = parseInt(digits.slice(4, 5), 10);

  if (!year || (semester !== 1 && semester !== 2)) return null;

  return Object.freeze({
    year: year,
    semester: semester
  });
}

/**
 * Faz o parse de um ID de semestre no formato YYYY/S.
 *
 * Exemplo:
 * "2026/1" -> { year: 2026, semester: 1 }
 *
 * @param {string} semesterId
 * @return {Object|null}
 */
function core_parseSemesterId_(semesterId) {
  const m = String(semesterId || "").trim().match(/^(\d{4})\/([12])$/);
  if (!m) return null;

  return Object.freeze({
    year: parseInt(m[1], 10),
    semester: parseInt(m[2], 10)
  });
}

/**
 * Calcula o semestre atual teórico do aluno com base no RGA
 * e no semestre institucional vigente.
 *
 * Regra:
 * semestreAtual = diferença de semestres + 1
 *
 * Exemplo:
 * entrada 2023/1
 * atual   2026/1
 * resultado = 7
 *
 * @param {string|number} rga
 * @param {Date=} refDate
 * @return {number|null}
 */
function core_getStudentCurrentSemesterFromRga_(rga, refDate) {
  const entry = core_parseEntrySemesterFromRga_(rga);
  if (!entry) return null;

  const currentSemester = core_getCurrentSemester_(refDate);
  if (!currentSemester || !currentSemester.id) return null;

  const current = core_parseSemesterId_(currentSemester.id);
  if (!current) return null;

  const diff =
    (current.year - entry.year) * 2 +
    (current.semester - entry.semester);

  const semesterNumber = diff + 1;

  if (semesterNumber < 1) return null;
  return semesterNumber;
}

/* ======================================================================
 * Semestre institucional por data / semestres concluídos no grupo
 * ====================================================================== */

/**
 * Retorna o semestre institucional correspondente a uma data.
 *
 * Regra:
 * 1. se a data cair dentro de um semestre, retorna esse semestre
 * 2. se cair em um buraco entre semestres, retorna o próximo semestre futuro
 *
 * @param {Date|string|number} refDate
 * @return {Object|null}
 */
function core_getSemesterForDate_(refDate) {
  const target = core_parseDateOrNull_(refDate);
  if (!target) return null;

  const data = core_getRowsFromSheetByKey_("VIGENCIA_SEMESTRES");
  const idx = core_getHeaderIndexMap_(data.headers, {
    semesterId: "ID_Semestre",
    start: "Início",
    end: "Fim",
    periodId: "ID_Período"
  });

  if (idx.semesterId < 0 || idx.start < 0 || idx.end < 0) {
    throw new Error("core_getSemesterForDate_: cabeçalhos obrigatórios não encontrados em VIGENCIA_SEMESTRES.");
  }

  let nextSemester = null;

  for (let i = 0; i < data.rows.length; i++) {
    const row = data.rows[i];
    const semesterId = String(row[idx.semesterId] || "").trim();
    const start = core_parseDateOrNull_(row[idx.start]);
    const end = core_parseDateOrNull_(row[idx.end]);
    const periodId = idx.periodId >= 0 ? String(row[idx.periodId] || "").trim() : "";

    if (!semesterId || !start) continue;

    if (core_isDateInRange_(target, start, end)) {
      return Object.freeze({
        id: semesterId,
        startDate: start,
        endDate: end,
        periodId: periodId
      });
    }

    if (start > target) {
      if (!nextSemester || start < nextSemester.startDate) {
        nextSemester = Object.freeze({
          id: semesterId,
          startDate: start,
          endDate: end,
          periodId: periodId
        });
      }
    }
  }

  return nextSemester || null;
}

function core_getSemesterIdForDate_(refDate) {
  var semester = core_getSemesterForDate_(refDate);
  return semester && semester.id ? semester.id : "";
}

/**
 * Retorna o último semestre institucional concluído na data informada.
 *
 * Regra:
 * - considera apenas semestres cujo Fim seja anterior ou igual à data de referência
 * - se nenhum semestre tiver sido concluído ainda, retorna null
 *
 * @param {Date=} refDate
 * @return {Object|null}
 */
function core_getLastCompletedSemester_(refDate) {
  const now = refDate || new Date();
  const data = core_getRowsFromSheetByKey_("VIGENCIA_SEMESTRES");
  const idx = core_getHeaderIndexMap_(data.headers, {
    semesterId: "ID_Semestre",
    start: "Início",
    end: "Fim",
    periodId: "ID_Período"
  });

  if (idx.semesterId < 0 || idx.end < 0) {
    throw new Error("core_getLastCompletedSemester_: cabeçalhos obrigatórios não encontrados em VIGENCIA_SEMESTRES.");
  }

  let lastCompleted = null;

  for (let i = 0; i < data.rows.length; i++) {
    const row = data.rows[i];
    const semesterId = String(row[idx.semesterId] || "").trim();
    const start = idx.start >= 0 ? core_parseDateOrNull_(row[idx.start]) : null;
    const end = core_parseDateOrNull_(row[idx.end]);
    const periodId = idx.periodId >= 0 ? String(row[idx.periodId] || "").trim() : "";

    if (!semesterId || !end) continue;
    if (end > now) continue;

    if (!lastCompleted || end > lastCompleted.endDate) {
      lastCompleted = Object.freeze({
        id: semesterId,
        startDate: start,
        endDate: end,
        periodId: periodId
      });
    }
  }

  return lastCompleted;
}

/**
 * Retorna a quantidade de semestres entre dois IDs institucionais.
 *
 * Exemplo:
 * de 2024/1 até 2025/2 -> 4
 *
 * @param {string} fromSemesterId
 * @param {string} toSemesterId
 * @return {number|null}
 */
function core_getSemesterCountBetweenIds_(fromSemesterId, toSemesterId) {
  const from = core_parseSemesterId_(fromSemesterId);
  const to = core_parseSemesterId_(toSemesterId);

  if (!from || !to) return null;

  const diff =
    (to.year - from.year) * 2 +
    (to.semester - from.semester);

  const count = diff + 1;
  return count < 1 ? null : count;
}

/**
 * Calcula o número de semestres concluídos no grupo a partir da data de entrada.
 *
 * Regra:
 * - identifica o semestre institucional correspondente à data de entrada
 * - identifica o último semestre concluído
 * - conta quantos semestres completos existem entre eles
 *
 * Exemplo:
 * entrada em 2024/1
 * último concluído = 2025/2
 * resultado = 4
 *
 * @param {Date|string|number} entryDate
 * @param {Date=} refDate
 * @return {number|null}
 */
function core_getCompletedGroupSemesterCountFromEntryDate_(entryDate, refDate) {
  const entrySemester = core_getSemesterForDate_(entryDate);
  if (!entrySemester || !entrySemester.id) return null;

  const lastCompleted = core_getLastCompletedSemester_(refDate);
  if (!lastCompleted || !lastCompleted.id) return null;

  return core_getSemesterCountBetweenIds_(entrySemester.id, lastCompleted.id);
}

/* ======================================================================
 * Semestre institucional por ID curto / semestres concluídos no grupo
 * ====================================================================== */

/**
 * Faz o parse de um ID de semestre nos formatos:
 * - YY/S   (ex.: 25/2)
 * - YYYY/S (ex.: 2025/2)
 *
 * @param {string|number} semesterIdRaw
 * @return {Object|null}
 */
function core_parseShortSemesterId_(semesterIdRaw) {
  const raw = String(semesterIdRaw || "").trim().replace(/\s+/g, "");

  // formato YYYY/S
  let m = raw.match(/^(\d{4})\/([12])$/);
  if (m) {
    return Object.freeze({
      year: parseInt(m[1], 10),
      semester: parseInt(m[2], 10)
    });
  }

  // formato YY/S
  m = raw.match(/^(\d{2})\/([12])$/);
  if (m) {
    return Object.freeze({
      year: 2000 + parseInt(m[1], 10),
      semester: parseInt(m[2], 10)
    });
  }

  return null;
}

/**
 * Retorna a quantidade de semestres entre um ID curto (YY/S)
 * e um ID completo (YYYY/S).
 *
 * Exemplo:
 * fromShort = "25/2"
 * toFull = "2025/2"
 * resultado = 1
 *
 * @param {string} fromShortSemesterId
 * @param {string} toFullSemesterId
 * @return {number|null}
 */
function core_getSemesterCountFromShortToFullId_(fromShortSemesterId, toFullSemesterId) {
  const from = core_parseShortSemesterId_(fromShortSemesterId);
  const to = core_parseSemesterId_(toFullSemesterId);

  if (!from || !to) return null;

  const diff =
    (to.year - from.year) * 2 +
    (to.semester - from.semester);

  const count = diff + 1;
  return count < 1 ? null : count;
}

/**
 * Calcula o número de semestres concluídos no grupo a partir
 * do semestre de entrada no grupo no formato YY/S.
 *
 * Exemplo:
 * entrada = "25/2"
 * último concluído = "2025/2"
 * resultado = 1
 *
 * @param {string} entrySemesterShort
 * @param {Date=} refDate
 * @return {number|null}
 */
function core_getCompletedGroupSemesterCountFromEntrySemester_(entrySemesterShort, refDate) {
  const lastCompleted = core_getLastCompletedSemester_(refDate);
  if (!lastCompleted || !lastCompleted.id) return null;

  return core_getSemesterCountFromShortToFullId_(entrySemesterShort, lastCompleted.id);
}
