/***************************************
 * 11_core_members.gs
 *
 * Camada de acesso aos dados dos membros do GEAPA.
 *
 * Objetivo:
 * - Centralizar a leitura da planilha de membros no CORE
 * - Evitar que cada módulo faça parsing manual da aba "Dados dos Membros"
 * - Permitir consultas reutilizáveis por cargo/função
 *
 * Padrão de uso:
 * - Fonte de dados via Registry (KEY da planilha/aba)
 * - Funções internas privadas com sufixo "_"
 * - Exportação pública feita separadamente em 20_public_exports.gs
 *
 * Observações:
 * - Este arquivo NÃO exporta funções públicas sozinho.
 * - Depois dele, ainda é preciso adicionar wrappers em:
 *   - 20_public_exports.gs
 *   - 00_core_public_api.gs
 ***************************************/

/**
 * Configuração da leitura da planilha de membros.
 *
 * Ajuste estes cabeçalhos se a estrutura real da planilha mudar.
 */
const CORE_MEMBERS_CFG = Object.freeze({
  /**
   * KEY do Registry que aponta para a planilha/aba de membros.
   * Exemplo esperado no Registry:
   * KEY = MEMBERS_ATUAIS
   */
  registryKey: "MEMBERS_ATUAIS",

  /**
   * Linha do cabeçalho.
   * Mantido explícito para facilitar manutenção futura.
   */
  headerRow: 1,

  /**
   * Nomes esperados dos cabeçalhos da planilha.
   * Se os nomes reais forem diferentes, altere aqui.
   */
  headers: Object.freeze({
    name: Object.freeze(["Membro", "MEMBRO", "NOME_MEMBRO", "Nome"]),
    role: Object.freeze(["Cargo/Fun\u00E7\u00E3o atual", "Cargo/fun\u00E7\u00E3o atual", "Cargo/funcao atual", "CARGO_FUNCAO_ATUAL"]),
    phone: Object.freeze(["Telefone", "TELEFONE"]),
    email: Object.freeze(["Email", "E-mail", "EMAIL"]),
    status: Object.freeze(["Status", "STATUS_CADASTRAL"]),
    rga: Object.freeze(["RGA"])
  }),

  /**
   * Valores considerados "ativos" quando existir coluna de status.
   * Comparação feita em minúsculas e sem espaços extras.
   */
  activeValues: Object.freeze(["ativo", "ativa", "ok"]),

  /**
   * Cargos estratégicos para consultas institucionais rápidas.
   */
  leadershipRoles: Object.freeze({
    presidente: "Presidente",
    vicePresidente: "Vice-presidente",
    secretarioGeral: "Secretário(a) Geral",
    secretarioExecutivo: "Secretário(a) Executivo(a)",
    diretorComunicacao: "Diretor(a) de Comunicação"
  })
});


/* ======================================================================
 * Helpers internos
 * ====================================================================== */

/**
 * Normaliza cabeçalhos/textos para comparação.
 *
 * Regras:
 * - remove espaços extras nas bordas
 * - converte para minúsculas
 *
 * @param {*} value
 * @return {string}
 */
function core_normalizeMemberText_(value) {
  return String(value || "").trim().toLowerCase();
}

/**
 * Lê a planilha de membros via Registry.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function core_getMembersSheet_() {
  return core_getSheetByKey_(CORE_MEMBERS_CFG.registryKey);
}

/**
 * Retorna um mapa de índices dos cabeçalhos encontrados.
 *
 * Exemplo de retorno:
 * {
 *   name: 0,
 *   role: 1,
 *   phone: 2,
 *   email: 3,
 *   status: 4,
 *   rga: 5
 * }
 *
 * Se algum cabeçalho não existir, o índice será -1.
 *
 * @param {string[]} headers
 * @return {Object}
 */
function core_findMemberHeaderIndex_(normalizedHeaders, aliases) {
  const names = Array.isArray(aliases) ? aliases : [aliases];

  for (let i = 0; i < names.length; i++) {
    const idx = normalizedHeaders.indexOf(core_normalizeMemberText_(names[i]));
    if (idx !== -1) return idx;
  }

  return -1;
}

function core_getMembersHeaderIndexMap_(headers) {
  const normalized = headers.map(core_normalizeMemberText_);

  return {
    name: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.name),
    role: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.role),
    phone: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.phone),
    email: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.email),
    status: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.status),
    rga: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.rga)
  };
}

/**
 * Verifica se uma linha representa um membro "ativo".
 *
 * Regra:
 * - se não existir coluna de status, considera ativo
 * - se a célula de status estiver vazia, considera ativo
 * - se existir valor, ele precisa estar em CORE_MEMBERS_CFG.activeValues
 *
 * @param {Array} row
 * @param {Object} idx
 * @return {boolean}
 */
function core_isActiveMemberRow_(row, idx) {
  if (idx.status < 0) return true;

  const status = core_normalizeMemberText_(row[idx.status]);
  if (!status) return true;

  return CORE_MEMBERS_CFG.activeValues.indexOf(status) !== -1;
}

/**
 * Converte uma linha da planilha em objeto padronizado de membro.
 *
 * @param {Array} row
 * @param {Object} idx
 * @return {Object}
 */
function core_mapMemberRow_(row, idx) {
  return Object.freeze({
    name: idx.name >= 0 ? String(row[idx.name] || "").trim() : "",
    role: idx.role >= 0 ? String(row[idx.role] || "").trim() : "",
    phone: idx.phone >= 0 ? String(row[idx.phone] || "").trim() : "",
    email: idx.email >= 0 ? String(row[idx.email] || "").trim() : "",
    status: idx.status >= 0 ? String(row[idx.status] || "").trim() : "",
    rga: idx.rga >= 0 ? String(row[idx.rga] || "").trim() : ""
  });
}


/* ======================================================================
 * Funções internas principais
 * ====================================================================== */

/**
 * Retorna todos os membros válidos da planilha.
 *
 * Critérios:
 * - ignora linhas sem nome
 * - aplica filtro de status quando a coluna existir
 *
 * @return {Object[]} Lista de membros padronizados
 */
function core_getMembers_() {
  const sh = core_getMembersSheet_();
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  if (!values || !values.length) return [];

  const headers = values.shift().map(v => String(v || "").trim());
  const idx = core_getMembersHeaderIndexMap_(headers);

  if (idx.name < 0) {
    throw new Error(
      'core_getMembers_: cabeçalho obrigatório não encontrado: "' +
      CORE_MEMBERS_CFG.headers.name +
      '".'
    );
  }

  return values
    .filter(row => {
      const name = String(row[idx.name] || "").trim();
      if (!name) return false;
      return core_isActiveMemberRow_(row, idx);
    })
    .map(row => core_mapMemberRow_(row, idx));
}

/**
 * Retorna membros filtrados por cargo/função.
 *
 * Comparação:
 * - case-insensitive
 * - trim automático
 *
 * @param {string} role
 * @return {Object[]}
 */
function core_getMembersByRole_(role) {
  const wanted = core_normalizeMemberText_(role);
  if (!wanted) return [];

  return core_getMembers_().filter(member => {
    return core_normalizeMemberText_(member.role) === wanted;
  });
}

/**
 * Retorna o primeiro membro encontrado para um cargo.
 *
 * Útil quando a expectativa institucional é haver apenas 1 ocupante.
 *
 * @param {string} role
 * @return {Object|null}
 */
function core_getFirstMemberByRole_(role) {
  const found = core_getMembersByRole_(role);
  return found.length ? found[0] : null;
}

/**
 * Retorna um objeto com cargos estratégicos da gestão.
 *
 * Estrutura:
 * {
 *   presidente: {...} | null,
 *   vicePresidente: {...} | null,
 *   secretarioGeral: {...} | null,
 *   secretarioExecutivo: {...} | null,
 *   diretorComunicacao: {...} | null
 * }
 *
 * @return {Object}
 */
function core_getLeadership_() {
  return Object.freeze({
    presidente: core_getFirstMemberByRole_(CORE_MEMBERS_CFG.leadershipRoles.presidente),
    vicePresidente: core_getFirstMemberByRole_(CORE_MEMBERS_CFG.leadershipRoles.vicePresidente),
    secretarioGeral: core_getFirstMemberByRole_(CORE_MEMBERS_CFG.leadershipRoles.secretarioGeral),
    secretarioExecutivo: core_getFirstMemberByRole_(CORE_MEMBERS_CFG.leadershipRoles.secretarioExecutivo),
    diretorComunicacao: core_getFirstMemberByRole_(CORE_MEMBERS_CFG.leadershipRoles.diretorComunicacao)
  });
}
