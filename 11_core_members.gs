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
    occupation: core_getOccupationHeaderAliases_('currentOccupation'),
    phone: Object.freeze(["Telefone", "TELEFONE"]),
    email: Object.freeze(["Email", "E-mail", "EMAIL"]),
    status: Object.freeze(["Status", "STATUS_CADASTRAL"]),
    rga: Object.freeze(["RGA"]),
    situacaoGeral: Object.freeze([
      "SITUACAO_GERAL",
      "SITUA\u00C7\u00C3O_GERAL",
      "Situacao geral",
      "Situa\u00E7\u00E3o geral",
      "Status geral",
      "Status",
      "STATUS_CADASTRAL"
    ]),
    vinculo: Object.freeze([
      "VINCULO",
      "V\u00CDNCULO",
      "Vinculo",
      "V\u00EDnculo",
      "STATUS_VINCULO",
      "TIPO_VINCULO"
    ])
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
  return String(value || "")
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
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
 *   occupation: 1,
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
    occupation: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.occupation),
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
function core_mapMemberRow_(row, idx, headers) {
  const occupationCell = headers
    ? core_getOccupationValueFromRowByHeaders_(row, headers, 'currentOccupation', core_normalizeMemberText_)
    : { found: idx.occupation >= 0, value: idx.occupation >= 0 ? row[idx.occupation] : '' };
  const occupation = occupationCell.found ? String(occupationCell.value || "").trim() : "";

  return Object.freeze({
    name: idx.name >= 0 ? String(row[idx.name] || "").trim() : "",
    occupation: occupation,
    role: occupation,
    phone: idx.phone >= 0 ? String(row[idx.phone] || "").trim() : "",
    email: idx.email >= 0 ? String(row[idx.email] || "").trim() : "",
    status: idx.status >= 0 ? String(row[idx.status] || "").trim() : "",
    rga: idx.rga >= 0 ? String(row[idx.rga] || "").trim() : ""
  });
}

function core_getPortalMemberHeaderIndexMap_(headers) {
  const normalized = headers.map(core_normalizeMemberText_);
  const occupationAliases = core_getOccupationHeaderAliases_('currentOccupation');

  return {
    name: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.name),
    email: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.email),
    rga: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.rga),
    status: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.status),
    occupation: core_findMemberHeaderIndex_(normalized, occupationAliases),
    situacaoGeral: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.situacaoGeral),
    vinculo: core_findMemberHeaderIndex_(
      normalized,
      CORE_MEMBERS_CFG.headers.vinculo.concat(occupationAliases)
    )
  };
}

function core_normalizePortalMemberLookup_(emailOuRga) {
  const raw = String(emailOuRga == null ? "" : emailOuRga).trim();
  if (!raw) {
    return Object.freeze({
      raw: "",
      email: "",
      rgaKey: ""
    });
  }

  if (raw.indexOf("@") >= 0) {
    return Object.freeze({
      raw: raw,
      email: core_extractEmailAddress_(raw),
      rgaKey: ""
    });
  }

  return Object.freeze({
    raw: raw,
    email: "",
    rgaKey: core_normalizeIdentityKey_(raw)
  });
}

function core_getPortalMemberCell_(row, idx, fallbackValue) {
  if (idx == null || idx < 0) return fallbackValue || "";
  const value = String(row[idx] || "").trim();
  return value || fallbackValue || "";
}

function core_mapPortalMemberRow_(row, idx, opts) {
  opts = opts || {};
  const requireValidEmail = opts.requireValidEmail !== false;
  const useDefaultLabels = opts.useDefaultLabels !== false;
  const emailCadastrado = idx.email >= 0 && core_isValidEmail_(row[idx.email])
    ? core_extractEmailAddress_(row[idx.email])
    : "";
  const rga = core_getPortalMemberCell_(row, idx.rga, "");
  const situacaoGeral =
    core_getPortalMemberCell_(row, idx.situacaoGeral, "") ||
    core_getPortalMemberCell_(row, idx.status, "") ||
    (useDefaultLabels ? "Ativo" : "");
  const vinculo =
    core_getPortalMemberCell_(row, idx.vinculo, "") ||
    core_getPortalMemberCell_(row, idx.occupation, "") ||
    (useDefaultLabels ? "Membro" : "");

  if (requireValidEmail && !emailCadastrado) {
    return null;
  }

  return Object.freeze({
    id: rga || "",
    nomeExibicao: core_getPortalMemberCell_(row, idx.name, ""),
    emailCadastrado: emailCadastrado,
    rga: rga,
    situacaoGeral: situacaoGeral,
    vinculo: vinculo
  });
}

function core_buscarMembroParaPortalInSheet_(sheet, emailOuRga, opts) {
  opts = opts || {};
  const lookup = core_normalizePortalMemberLookup_(emailOuRga);
  if (!lookup.raw) return null;
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= CORE_MEMBERS_CFG.headerRow || lastCol < 1) return null;

  const headers = sheet
    .getRange(CORE_MEMBERS_CFG.headerRow, 1, 1, lastCol)
    .getValues()[0]
    .map(function(header) {
      return String(header || "").trim();
    });
  const idx = core_getPortalMemberHeaderIndexMap_(headers);

  if (idx.email < 0 || idx.rga < 0) {
    throw new Error("Schema de MEMBERS_ATUAIS invalido para consulta do Portal.");
  }

  const startRow = CORE_MEMBERS_CFG.headerRow + 1;
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowEmail = idx.email >= 0 ? core_extractEmailAddress_(row[idx.email]) : "";
    const rowRgaKey = idx.rga >= 0 ? core_normalizeIdentityKey_(row[idx.rga]) : "";
    const foundByEmail = lookup.email && rowEmail && lookup.email === rowEmail;
    const foundByRga = lookup.rgaKey && rowRgaKey && lookup.rgaKey === rowRgaKey;

    if (!foundByEmail && !foundByRga) continue;

    return core_mapPortalMemberRow_(row, idx, opts);
  }

  return null;
}

/**
 * Contrato inicial com o geapa-portal.
 *
 * O backend Apps Script do portal usa esta consulta para localizar um unico
 * membro por e-mail ou RGA e decidir o e-mail cadastrado que recebera o codigo
 * de acesso. O navegador nunca deve chamar esta funcao diretamente.
 *
 * Retorna somente:
 * id, nomeExibicao, emailCadastrado, rga, situacaoGeral, vinculo.
 *
 * @param {string} emailOuRga
 * @return {Object|null}
 */
function core_buscarMembroParaPortal_(emailOuRga) {
  return core_buscarMembroParaPortalInSheet_(core_getMembersSheet_(), emailOuRga);
}

function core_buildPortalError_(code, message) {
  return Object.freeze({
    ok: false,
    code: String(code || "ERRO_PORTAL").trim(),
    message: String(message || "Nao foi possivel concluir a consulta.").trim()
  });
}

function core_buildMinhaSituacaoPortalVazia_() {
  return core_buildMinhaSituacaoPortal_(Object.freeze([]));
}

function core_buildMinhaSituacaoPortal_(pendencias) {
  const pending = Object.freeze((pendencias || []).slice());

  return Object.freeze({
    resumo: Object.freeze({
      frequencia: "",
      pendenciasAbertas: pending.length,
      certificadosDisponiveis: 0
    }),
    pendencias: pending,
    participacao: Object.freeze({
      frequenciaGeral: "",
      atividadesRecentes: Object.freeze([])
    }),
    certificados: Object.freeze([]),
    avisos: Object.freeze([])
  });
}

function core_isPortalUndefinedValue_(value) {
  const normalized = core_normalizeMemberText_(value);
  if (!normalized) return true;

  return [
    "indefinido",
    "indefinida",
    "nao informado",
    "nao informada",
    "nao consta",
    "sem informacao",
    "a definir"
  ].indexOf(normalized) !== -1;
}

function core_buildPortalPendencia_(tipo, titulo, descricao, severidade) {
  return Object.freeze({
    tipo: tipo,
    titulo: titulo,
    descricao: descricao,
    severidade: severidade,
    status: "pendente"
  });
}

function core_getPortalPendenciasCadastro_(membro) {
  const pendencias = [];

  if (!membro.emailCadastrado || !core_isValidEmail_(membro.emailCadastrado)) {
    pendencias.push(core_buildPortalPendencia_(
      "cadastro",
      "E-mail cadastrado ausente ou invalido",
      "Procure a Diretoria para atualizar seu e-mail de contato no cadastro do GEAPA.",
      "alta"
    ));
  }

  if (!String(membro.rga || "").trim()) {
    pendencias.push(core_buildPortalPendencia_(
      "cadastro",
      "RGA nao informado",
      "Procure a Diretoria para atualizar seu RGA no cadastro do GEAPA.",
      "media"
    ));
  }

  if (!String(membro.nomeExibicao || "").trim()) {
    pendencias.push(core_buildPortalPendencia_(
      "cadastro",
      "Nome de exibicao nao informado",
      "Procure a Diretoria para atualizar seu nome no cadastro do GEAPA.",
      "media"
    ));
  }

  if (core_isPortalUndefinedValue_(membro.vinculo)) {
    pendencias.push(core_buildPortalPendencia_(
      "cadastro",
      "Vinculo cadastral indefinido",
      "Procure a Diretoria para confirmar seu vinculo cadastral no GEAPA.",
      "baixa"
    ));
  }

  if (core_isPortalUndefinedValue_(membro.situacaoGeral)) {
    pendencias.push(core_buildPortalPendencia_(
      "administrativo",
      "Situacao geral indefinida",
      "Procure a Diretoria para confirmar sua situacao cadastral no GEAPA.",
      "baixa"
    ));
  }

  return Object.freeze(pendencias);
}

function core_buildMinhaSituacaoPortalResponse_(membro) {
  const pendencias = core_getPortalPendenciasCadastro_(membro);

  return Object.freeze({
    ok: true,
    membro: Object.freeze({
      id: String(membro.id || "").trim(),
      nomeExibicao: String(membro.nomeExibicao || "").trim(),
      emailCadastrado: String(membro.emailCadastrado || "").trim(),
      rga: String(membro.rga || "").trim(),
      vinculo: String(membro.vinculo || "").trim(),
      situacaoGeral: String(membro.situacaoGeral || "").trim()
    }),
    minhaSituacao: core_buildMinhaSituacaoPortal_(pendencias)
  });
}

function core_buscarMinhaSituacaoParaPortalInSheet_(sheet, emailOuRga) {
  const membro = core_buscarMembroParaPortalInSheet_(sheet, emailOuRga, {
    requireValidEmail: false,
    useDefaultLabels: false
  });

  if (!membro) {
    return core_buildPortalError_(
      "MEMBRO_NAO_ENCONTRADO",
      "Membro nao encontrado para o e-mail ou RGA informado."
    );
  }

  return core_buildMinhaSituacaoPortalResponse_(membro);
}

/**
 * Contrato inicial da tela "Minha situacao" do geapa-portal.
 *
 * Esta V1 usa somente fontes oficiais ja centralizadas no Core para localizar
 * o proprio membro. As pendencias retornadas sao apenas cadastrais objetivas
 * derivadas desse proprio registro. Enquanto frequencia, certificados e
 * atividades recentes nao tiverem uma fonte confiavel integrada ao Core, esses
 * blocos ficam vazios ou zerados. Nao inventar dados nesta funcao.
 *
 * Regras de seguranca para futuras diretorias:
 * - retornar apenas dados do membro localizado;
 * - nunca retornar listas completas de membros;
 * - nao expor IDs internos de planilhas, tokens, chaves ou dados de terceiros;
 * - manter a filtragem no backend Apps Script, nunca no front-end.
 *
 * @param {string} emailOuRga
 * @return {Object}
 */
function core_buscarMinhaSituacaoParaPortal_(emailOuRga) {
  return core_buscarMinhaSituacaoParaPortalInSheet_(core_getMembersSheet_(), emailOuRga);
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
    .map(row => core_mapMemberRow_(row, idx, headers));
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
function core_getMembersByOccupation_(occupation) {
  const wanted = core_normalizeMemberText_(occupation);
  if (!wanted) return [];

  return core_getMembers_().filter(member => {
    return core_normalizeMemberText_(member.occupation) === wanted;
  });
}

function core_getMembersByRole_(role) {
  return core_getMembersByOccupation_(role);
}

/**
 * Retorna o primeiro membro encontrado para um cargo.
 *
 * Útil quando a expectativa institucional é haver apenas 1 ocupante.
 *
 * @param {string} role
 * @return {Object|null}
 */
function core_getFirstMemberByOccupation_(occupation) {
  const found = core_getMembersByOccupation_(occupation);
  return found.length ? found[0] : null;
}

function core_getFirstMemberByRole_(role) {
  return core_getFirstMemberByOccupation_(role);
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
    presidente: core_getFirstMemberByOccupation_(CORE_MEMBERS_CFG.leadershipRoles.presidente),
    vicePresidente: core_getFirstMemberByOccupation_(CORE_MEMBERS_CFG.leadershipRoles.vicePresidente),
    secretarioGeral: core_getFirstMemberByOccupation_(CORE_MEMBERS_CFG.leadershipRoles.secretarioGeral),
    secretarioExecutivo: core_getFirstMemberByOccupation_(CORE_MEMBERS_CFG.leadershipRoles.secretarioExecutivo),
    diretorComunicacao: core_getFirstMemberByOccupation_(CORE_MEMBERS_CFG.leadershipRoles.diretorComunicacao)
  });
}
