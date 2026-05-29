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
    ]),
    periodoUltimaApresentacao: Object.freeze(["PERIODO_ULTIMA_APRESENTACAO"]),
    qtdApresentacoesRealizadas: Object.freeze(["QTD_APRESENTACOES_REALIZADAS"]),
    qtdDiasQueContamParaLimiteDiretoria: Object.freeze(["QTD_DIAS_QUE_CONTAM_PARA_LIMITE_DIRETORIA"]),
    limiteDiasDiretoria: Object.freeze(["LIMITE_DIAS_DIRETORIA"]),
    saldoDiasDiretoria: Object.freeze(["SALDO_DIAS_DIRETORIA"]),
    statusElegibilidadeDiretoria: Object.freeze(["STATUS_ELEGIBILIDADE_DIRETORIA"]),
    dataLimiteEstimadaDiretoria: Object.freeze(["DATA_LIMITE_ESTIMADA_DIRETORIA"]),
    dataIntegracao: Object.freeze([
      "DATA_INTEGRACAO",
      "Data integracao",
      "Data integração",
      "DATA_INGRESSO",
      "Data ingresso",
      "Data de ingresso",
      "INGRESSO"
    ]),
    dataDesligamento: Object.freeze([
      "DATA_DESLIGAMENTO",
      "Data desligamento",
      "Data de desligamento",
      "DATA_SAIDA",
      "Data saida",
      "Data saída",
      "SAIDA",
      "Saída"
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
    ),
    periodoUltimaApresentacao: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.periodoUltimaApresentacao),
    qtdApresentacoesRealizadas: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.qtdApresentacoesRealizadas),
    qtdDiasQueContamParaLimiteDiretoria: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.qtdDiasQueContamParaLimiteDiretoria),
    limiteDiasDiretoria: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.limiteDiasDiretoria),
    saldoDiasDiretoria: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.saldoDiasDiretoria),
    statusElegibilidadeDiretoria: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.statusElegibilidadeDiretoria),
    dataLimiteEstimadaDiretoria: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.dataLimiteEstimadaDiretoria)
  };
}

function core_getAttendanceMemberHeaderIndexMap_(headers) {
  const normalized = headers.map(core_normalizeMemberText_);
  const occupationAliases = core_getOccupationHeaderAliases_('currentOccupation');

  return {
    name: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.name),
    rga: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.rga),
    status: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.status),
    situacaoGeral: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.situacaoGeral),
    vinculo: core_findMemberHeaderIndex_(
      normalized,
      CORE_MEMBERS_CFG.headers.vinculo.concat(occupationAliases)
    ),
    dataIntegracao: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.dataIntegracao),
    dataDesligamento: core_findMemberHeaderIndex_(normalized, CORE_MEMBERS_CFG.headers.dataDesligamento)
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

function core_parsePortalNonNegativeNumber_(value) {
  if (value == null || String(value).trim() === "") return 0;
  const normalized = String(value).trim().replace(",", ".");
  const parsed = Number(normalized);
  if (!isFinite(parsed) || isNaN(parsed) || parsed < 0) return 0;
  return parsed;
}

function core_buildPortalApresentacoesVazio_() {
  return Object.freeze({
    periodoUltimaApresentacao: "",
    quantidadeRealizadas: 0
  });
}

function core_buildPortalApresentacoesFromRow_(row, idx) {
  return Object.freeze({
    periodoUltimaApresentacao: core_getPortalMemberCell_(row, idx.periodoUltimaApresentacao, ""),
    quantidadeRealizadas: core_parsePortalNonNegativeNumber_(
      core_getPortalMemberCell_(row, idx.qtdApresentacoesRealizadas, "")
    )
  });
}

function core_buildPortalDiretoriaVazio_() {
  return Object.freeze({
    statusElegibilidade: "",
    diasComputados: 0,
    limiteDias: 0,
    saldoDias: 0,
    dataLimiteEstimada: ""
  });
}

function core_buildPortalDiretoriaFromRow_(row, idx) {
  return Object.freeze({
    statusElegibilidade: core_getPortalMemberCell_(row, idx.statusElegibilidadeDiretoria, ""),
    diasComputados: core_parsePortalNonNegativeNumber_(
      core_getPortalMemberCell_(row, idx.qtdDiasQueContamParaLimiteDiretoria, "")
    ),
    limiteDias: core_parsePortalNonNegativeNumber_(
      core_getPortalMemberCell_(row, idx.limiteDiasDiretoria, "")
    ),
    saldoDias: core_parsePortalNonNegativeNumber_(
      core_getPortalMemberCell_(row, idx.saldoDiasDiretoria, "")
    ),
    dataLimiteEstimada: core_getPortalMemberCell_(row, idx.dataLimiteEstimadaDiretoria, "")
  });
}

function core_mapPortalMemberRow_(row, idx, opts) {
  opts = opts || {};
  const requireValidEmail = opts.requireValidEmail !== false;
  const useDefaultLabels = opts.useDefaultLabels !== false;
  const includePresentationData = opts.includePresentationData === true;
  const includeBoardEligibilityData = opts.includeBoardEligibilityData === true;
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

  const member = {
    id: rga || "",
    nomeExibicao: core_getPortalMemberCell_(row, idx.name, ""),
    emailCadastrado: emailCadastrado,
    rga: rga,
    situacaoGeral: situacaoGeral,
    vinculo: vinculo
  };

  if (includePresentationData) {
    member._portalApresentacoes = core_buildPortalApresentacoesFromRow_(row, idx);
  }

  if (includeBoardEligibilityData) {
    member._portalDiretoria = core_buildPortalDiretoriaFromRow_(row, idx);
  }

  return Object.freeze(member);
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
  const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol);
  const values = range.getValues();
  const displayValues = typeof range.getDisplayValues === "function"
    ? range.getDisplayValues()
    : values;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowEmail = idx.email >= 0 ? core_extractEmailAddress_(row[idx.email]) : "";
    const rowRgaKey = idx.rga >= 0 ? core_normalizeIdentityKey_(row[idx.rga]) : "";
    const foundByEmail = lookup.email && rowEmail && lookup.email === rowEmail;
    const foundByRga = lookup.rgaKey && rowRgaKey && lookup.rgaKey === rowRgaKey;

    if (!foundByEmail && !foundByRga) continue;

    return core_mapPortalMemberRow_(displayValues[i] || row, idx, opts);
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
  return core_buildMinhaSituacaoPortal_(
    Object.freeze([]),
    core_buildPortalApresentacoesVazio_(),
    core_buildPortalDiretoriaVazio_()
  );
}

function core_buildMinhaSituacaoPortal_(pendencias, apresentacoes, diretoria) {
  const pending = Object.freeze((pendencias || []).slice());
  const presentationSummary = apresentacoes || core_buildPortalApresentacoesVazio_();
  const boardEligibility = diretoria || core_buildPortalDiretoriaVazio_();

  return Object.freeze({
    resumo: Object.freeze({
      frequencia: "",
      pendenciasAbertas: pending.length,
      certificadosDisponiveis: 0
    }),
    pendencias: pending,
    participacao: Object.freeze({
      frequenciaGeral: "",
      atividadesRecentes: Object.freeze([]),
      apresentacoes: Object.freeze({
        periodoUltimaApresentacao: String(presentationSummary.periodoUltimaApresentacao || "").trim(),
        quantidadeRealizadas: core_parsePortalNonNegativeNumber_(presentationSummary.quantidadeRealizadas)
      })
    }),
    diretoria: Object.freeze({
      statusElegibilidade: String(boardEligibility.statusElegibilidade || "").trim(),
      diasComputados: core_parsePortalNonNegativeNumber_(boardEligibility.diasComputados),
      limiteDias: core_parsePortalNonNegativeNumber_(boardEligibility.limiteDias),
      saldoDias: core_parsePortalNonNegativeNumber_(boardEligibility.saldoDias),
      dataLimiteEstimada: String(boardEligibility.dataLimiteEstimada || "").trim()
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

function core_buildMinhaSituacaoPortalResponse_(membro, usuario) {
  const pendencias = core_getPortalPendenciasCadastro_(membro);
  const apresentacoes = membro._portalApresentacoes || core_buildPortalApresentacoesVazio_();
  const diretoria = membro._portalDiretoria || core_buildPortalDiretoriaVazio_();

  const response = {
    ok: true,
    membro: Object.freeze({
      id: String(membro.id || "").trim(),
      nomeExibicao: String(membro.nomeExibicao || "").trim(),
      emailCadastrado: String(membro.emailCadastrado || "").trim(),
      rga: String(membro.rga || "").trim(),
      vinculo: String(membro.vinculo || "").trim(),
      situacaoGeral: String(membro.situacaoGeral || "").trim()
    }),
    minhaSituacao: core_buildMinhaSituacaoPortal_(pendencias, apresentacoes, diretoria)
  };

  if (usuario) {
    response.usuario = usuario;
  }

  return Object.freeze(response);
}

function core_buscarMinhaSituacaoParaPortalInSheet_(sheet, emailOuRga) {
  const membro = core_buscarMembroParaPortalInSheet_(sheet, emailOuRga, {
    requireValidEmail: false,
    useDefaultLabels: false,
    includePresentationData: true,
    includeBoardEligibilityData: true
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
  const membro = core_buscarMembroParaPortalInSheet_(core_getMembersSheet_(), emailOuRga, {
    requireValidEmail: false,
    useDefaultLabels: false,
    includePresentationData: true,
    includeBoardEligibilityData: true
  });

  if (!membro) {
    return core_buildPortalError_(
      "MEMBRO_NAO_ENCONTRADO",
      "Membro nao encontrado para o e-mail ou RGA informado."
    );
  }

  const usuarioResult = core_buscarUsuarioPortalFromMember_(membro);
  const usuario = usuarioResult && usuarioResult.ok ? usuarioResult.usuario : null;

  return core_buildMinhaSituacaoPortalResponse_(membro, usuario);
}

function core_normalizePortalCargoKeyFallback_(value) {
  const normalized = String(value == null ? "" : value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\(a\)/gi, "")
    .replace(/[()]/g, " ")
    .toUpperCase();

  const stopwords = {
    A: true,
    O: true,
    AS: true,
    OS: true,
    DE: true,
    DA: true,
    DO: true,
    DAS: true,
    DOS: true
  };

  return normalized
    .split(/[^A-Z0-9]+/)
    .filter(function(part) {
      return part && !stopwords[part];
    })
    .join("_");
}

function core_resolvePortalCargoKey_(assignment) {
  if (!assignment) return "";
  if (assignment.roleKey) return core_normalizePortalCargoKeyFallback_(assignment.roleKey);

  const candidates = [
    assignment.publicName,
    assignment.occupationName,
    assignment.occupation,
    assignment.rawOccupation,
    assignment.rawRole
  ];

  for (let i = 0; i < candidates.length; i++) {
    const text = String(candidates[i] || "").trim();
    if (!text) continue;

    try {
      const roleConfig = core_findInstitutionalRoleByAnyName_(text);
      if (roleConfig && roleConfig.roleKey) {
        return core_normalizePortalCargoKeyFallback_(roleConfig.roleKey);
      }
    } catch (err) {
      // Fallback textual abaixo preserva o contrato mesmo sem config carregada.
    }

    const fallback = core_normalizePortalCargoKeyFallback_(text);
    if (fallback) return fallback;
  }

  return "";
}

function core_resolvePortalCargoNome_(assignment) {
  if (!assignment) return "";
  return String(
    assignment.publicName ||
    assignment.occupationName ||
    assignment.occupation ||
    assignment.rawOccupation ||
    assignment.rawRole ||
    ""
  ).trim();
}

function core_resolvePortalGrupoCargo_(assignment, cargoKey) {
  if (assignment && assignment.roleConfig && assignment.roleConfig.group) {
    return core_normalizePortalCargoKeyFallback_(assignment.roleConfig.group);
  }

  if (cargoKey === "PRESIDENTE" || cargoKey === "VICE_PRESIDENTE") return "PRESIDENCIA";
  if (cargoKey === "SECRETARIO_GERAL" || cargoKey === "SECRETARIO_EXECUTIVO") return "SECRETARIA";
  if (cargoKey === "DIRETOR_COMUNICACAO" || cargoKey === "ASSESSOR_COMUNICACAO") return "COMUNICACAO";
  if (cargoKey === "CONSELHEIRO_CONSULTIVO") return "CONSELHO";

  const sourceType = core_normalizePortalCargoKeyFallback_(assignment && assignment.sourceType);
  if (sourceType === "DIRETORIA") return "DIRETORIA";
  if (sourceType === "ASSESSORIA") return "ASSESSORIA";
  if (sourceType === "CONSELHO") return "CONSELHO";

  return "";
}

function core_resolvePortalCargoFonte_(assignment) {
  const sourceType = core_normalizePortalCargoKeyFallback_(assignment && assignment.sourceType);
  if (sourceType === "DIRETORIA") return "VIGENCIAS_DIRETORES";
  if (sourceType === "ASSESSORIA") return "VIGENCIAS_ASSESSORES";
  if (sourceType === "CONSELHO") return "VIGENCIAS_CONSELHEIROS";
  return sourceType || "";
}

function core_formatPortalDateIso_(value) {
  const date = core_parseDateOrNull_(value);
  if (!date) return "";
  return core_formatDate_(date, null, "yyyy-MM-dd");
}

function core_buildPortalCargoAtual_(assignment) {
  const cargoKey = core_resolvePortalCargoKey_(assignment);
  return Object.freeze({
    cargoKey: cargoKey,
    cargoNome: core_resolvePortalCargoNome_(assignment),
    grupoCargo: core_resolvePortalGrupoCargo_(assignment, cargoKey),
    fonte: core_resolvePortalCargoFonte_(assignment),
    idDiretoria: String((assignment && (assignment.idDiretoria || assignment.boardId)) || "").trim(),
    dataInicio: core_formatPortalDateIso_(assignment && assignment.startDate),
    dataFimPrevista: core_formatPortalDateIso_(
      assignment && (assignment.endDatePlanned || assignment.endDateEffective || assignment.endDateReal)
    )
  });
}

function core_assignmentMatchesPortalMember_(assignment, membro) {
  if (!assignment || !membro) return false;

  const targetRga = core_roleProjectionNormalizeKey_(membro.rga || "");
  const targetEmail = core_roleProjectionNormalizeKey_(membro.emailCadastrado || "");
  const targetName = core_roleProjectionNormalizeKey_(membro.nomeExibicao || "");
  const rowRga = assignment.rgaNorm || core_roleProjectionNormalizeKey_(assignment.rga || "");
  const rowEmail = assignment.emailNorm || core_roleProjectionNormalizeKey_(assignment.email || "");
  const rowName = assignment.memberNameNorm || core_roleProjectionNormalizeKey_(assignment.memberName || "");

  if (targetRga && rowRga && targetRga === rowRga) return true;
  if (targetEmail && rowEmail && targetEmail === rowEmail) return true;
  if (!rowRga && !rowEmail && targetName && rowName && targetName === rowName) return true;

  return false;
}

function core_getPortalAssignmentsForMember_(membro, refDate) {
  const assignments = core_getCurrentInstitutionalAssignments_(refDate);
  return Object.freeze(assignments.filter(function(assignment) {
    return core_assignmentMatchesPortalMember_(assignment, membro);
  }));
}

function core_pushPortalProfile_(profiles, profile) {
  if (profiles.indexOf(profile) < 0) profiles.push(profile);
}

function core_derivePortalProfiles_(assignments) {
  const profiles = ["MEMBRO"];

  (assignments || []).forEach(function(assignment) {
    const cargoKey = core_resolvePortalCargoKey_(assignment);
    const sourceType = core_normalizePortalCargoKeyFallback_(assignment && assignment.sourceType);

    if (sourceType === "DIRETORIA") core_pushPortalProfile_(profiles, "DIRETORIA");
    if (sourceType === "ASSESSORIA") core_pushPortalProfile_(profiles, "ASSESSORIA");
    if (sourceType === "CONSELHO" || cargoKey === "CONSELHEIRO_CONSULTIVO") {
      core_pushPortalProfile_(profiles, "CONSELHO");
    }

    if (cargoKey === "PRESIDENTE" || cargoKey === "VICE_PRESIDENTE") {
      core_pushPortalProfile_(profiles, "PRESIDENCIA");
    }

    if (cargoKey === "SECRETARIO_GERAL" || cargoKey === "SECRETARIO_EXECUTIVO") {
      core_pushPortalProfile_(profiles, "SECRETARIA");
    }

    if (
      cargoKey === "DIRETOR_COMUNICACAO" ||
      cargoKey === "ASSESSOR_COMUNICACAO" ||
      cargoKey.indexOf("COMUNICACAO") >= 0
    ) {
      core_pushPortalProfile_(profiles, "COMUNICACAO");
    }
  });

  return Object.freeze(profiles);
}

function core_resolvePortalPerfilPrincipal_(profiles) {
  const priority = [
    "PRESIDENCIA",
    "DIRETORIA",
    "SECRETARIA",
    "COMUNICACAO",
    "CONSELHO",
    "ASSESSORIA",
    "MEMBRO"
  ];

  for (let i = 0; i < priority.length; i++) {
    if (profiles.indexOf(priority[i]) >= 0) return priority[i];
  }

  return "MEMBRO";
}

function core_derivePortalPermissions_(profiles) {
  const permissions = {
    podeVerAreaDiretoria: false,
    podeGerenciarAtividades: false,
    podeRegistrarChamada: false,
    podeEditarAtividade: false,
    podeAnalisarJustificativas: false,
    podeGerenciarCertificados: false,
    podeGerenciarComunicacao: false,
    podeGerenciarConfiguracoes: false
  };

  if (profiles.indexOf("DIRETORIA") >= 0) {
    permissions.podeVerAreaDiretoria = true;
    permissions.podeGerenciarAtividades = true;
    permissions.podeRegistrarChamada = true;
    permissions.podeEditarAtividade = true;
    permissions.podeAnalisarJustificativas = true;
  }

  if (profiles.indexOf("SECRETARIA") >= 0) {
    permissions.podeVerAreaDiretoria = true;
    permissions.podeGerenciarAtividades = true;
    permissions.podeRegistrarChamada = true;
    permissions.podeAnalisarJustificativas = true;
  }

  if (profiles.indexOf("PRESIDENCIA") >= 0) {
    permissions.podeVerAreaDiretoria = true;
    permissions.podeGerenciarAtividades = true;
  }

  if (profiles.indexOf("COMUNICACAO") >= 0) {
    permissions.podeGerenciarComunicacao = true;
  }

  return Object.freeze(permissions);
}

function core_buildUsuarioPortal_(membro, assignments) {
  const safeAssignments = assignments || [];
  const profiles = core_derivePortalProfiles_(safeAssignments);

  return Object.freeze({
    id: String((membro && (membro.id || membro.rga)) || "").trim(),
    nomeExibicao: String((membro && membro.nomeExibicao) || "").trim(),
    rga: String((membro && membro.rga) || "").trim(),
    emailCadastrado: String((membro && membro.emailCadastrado) || "").trim(),
    perfilPrincipal: core_resolvePortalPerfilPrincipal_(profiles),
    perfis: profiles,
    cargosAtuais: Object.freeze(safeAssignments.map(core_buildPortalCargoAtual_)),
    permissoes: core_derivePortalPermissions_(profiles)
  });
}

function core_buscarUsuarioPortalFromMember_(membro, refDate) {
  const assignments = core_getPortalAssignmentsForMember_(membro, refDate);
  return Object.freeze({
    ok: true,
    usuario: core_buildUsuarioPortal_(membro, assignments)
  });
}

function core_buscarUsuarioPortal_(emailOuRga) {
  const membro = core_buscarMembroParaPortal_(emailOuRga);

  if (!membro) {
    return core_buildPortalError_(
      "MEMBRO_NAO_ENCONTRADO",
      "Membro nao encontrado para o e-mail ou RGA informado."
    );
  }

  return core_buscarUsuarioPortalFromMember_(membro);
}

function core_buildUsuarioPortalTesteResumo_(result) {
  if (!result || !result.ok || !result.usuario) return result;

  const usuario = result.usuario;
  return Object.freeze({
    ok: true,
    usuario: Object.freeze({
      id: String(usuario.id || "").trim(),
      nomeExibicao: String(usuario.nomeExibicao || "").trim(),
      rga: String(usuario.rga || "").trim(),
      perfilPrincipal: String(usuario.perfilPrincipal || "").trim(),
      perfis: Object.freeze((usuario.perfis || []).slice()),
      cargosAtuais: Object.freeze((usuario.cargosAtuais || []).map(function(cargo) {
        return Object.freeze({
          cargoKey: cargo.cargoKey,
          cargoNome: cargo.cargoNome,
          grupoCargo: cargo.grupoCargo,
          fonte: cargo.fonte,
          idDiretoria: cargo.idDiretoria,
          dataInicio: cargo.dataInicio,
          dataFimPrevista: cargo.dataFimPrevista
        });
      })),
      permissoes: usuario.permissoes
    })
  });
}

function core_buildChamadaError_(errorCode, message) {
  return Object.freeze({
    ok: false,
    errorCode: String(errorCode || "ERRO_CHAMADA").trim(),
    message: String(message || "Nao foi possivel listar membros para chamada.").trim()
  });
}

function core_parseChamadaReferenceDate_(value) {
  const raw = String(value == null ? "" : value).trim();
  if (!raw) return null;

  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) {
    const year = Number(iso[1]);
    const month = Number(iso[2]);
    const day = Number(iso[3]);
    const date = new Date(year, month - 1, day);
    if (
      date.getFullYear() === year &&
      date.getMonth() === month - 1 &&
      date.getDate() === day
    ) {
      return core_startOfDay_(date);
    }
    return null;
  }

  const parsed = core_parseDateOrNull_(value);
  return parsed ? core_startOfDay_(parsed) : null;
}

function core_formatChamadaDateIso_(date) {
  return core_formatDate_(date, null, "yyyy-MM-dd");
}

function core_normalizeChamadaToken_(value) {
  return String(value == null ? "" : value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function core_isChamadaAuthorizedContext_(contexto) {
  if (!contexto || typeof contexto !== "object") return true;

  const perfil = core_normalizeChamadaToken_(
    contexto.perfil || contexto.perfilPrincipal || contexto.role || ""
  );
  if (!perfil) return true;

  return [
    "DIRETORIA",
    "PRESIDENCIA",
    "SECRETARIA",
    "SECRETARIO",
    "ADMIN_TECNICO"
  ].indexOf(perfil) >= 0;
}

function core_getChamadaCell_(row, idx, fallbackValue) {
  if (idx == null || idx < 0) return fallbackValue || "";
  const value = String(row[idx] || "").trim();
  return value || fallbackValue || "";
}

function core_parseChamadaRowDate_(row, idx) {
  if (idx == null || idx < 0) return null;
  return core_parseChamadaReferenceDate_(row[idx]);
}

function core_normalizeChamadaSituacao_(value) {
  const normalized = core_normalizeChamadaToken_(value);
  if (!normalized) return "";

  if (["ATIVO", "ATIVA", "OK", "REGULAR"].indexOf(normalized) >= 0) return "ATIVO";
  if (normalized.indexOf("SUSPENS") >= 0) return "SUSPENSO";
  if (normalized.indexOf("LICEN") >= 0) return "LICENCA";
  if (normalized.indexOf("AFAST") >= 0) return "AFASTADO";
  if (normalized.indexOf("DESLIG") >= 0) return "DESLIGADO";
  if (normalized.indexOf("INATIV") >= 0) return "INATIVO";

  return normalized.replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

function core_getChamadaCountingRuleFromSituation_(situacao) {
  const normalized = core_normalizeChamadaSituacao_(situacao);

  if (["SUSPENSO", "LICENCA", "AFASTADO"].indexOf(normalized) >= 0) {
    return Object.freeze({
      contaPresenca: false,
      contaFalta: false,
      motivoNaoAplicavel: "SITUACAO_NAO_CONTABILIZA_CHAMADA"
    });
  }

  if (["DESLIGADO", "INATIVO"].indexOf(normalized) >= 0) {
    return Object.freeze({
      contaPresenca: false,
      contaFalta: false,
      motivoNaoAplicavel: "VINCULO_NAO_ATIVO_NA_DATA"
    });
  }

  return Object.freeze({
    contaPresenca: true,
    contaFalta: true,
    motivoNaoAplicavel: ""
  });
}

function core_isChamadaLifecycleStatusUsable_(status) {
  const normalized = core_memberLifecycleNormalizeStatus_(status);
  return normalized && normalized !== "CANCELADO";
}

function core_getChamadaLifecycleEventsByRga_(observacoes) {
  const byRga = {};

  try {
    core_memberLifecycleListEvents_({
      excludeCanceled: true
    }).forEach(function(event) {
      if (!event || !event.rga || !event.eventDate) return;
      if (!core_isChamadaLifecycleStatusUsable_(event.eventStatus)) return;

      const key = core_normalizeIdentityKey_(event.rga);
      if (!key) return;
      if (!byRga[key]) byRga[key] = [];
      byRga[key].push(event);
    });
  } catch (err) {
    observacoes.push(
      "Historico MEMBER_EVENTOS_VINCULO indisponivel; aplicabilidade usa apenas campos atuais de MEMBERS_ATUAIS."
    );
  }

  Object.keys(byRga).forEach(function(key) {
    byRga[key].sort(function(a, b) {
      const aTime = a.eventDate ? a.eventDate.getTime() : 0;
      const bTime = b.eventDate ? b.eventDate.getTime() : 0;
      if (aTime !== bTime) return aTime - bTime;
      return String(a.eventId || "").localeCompare(String(b.eventId || ""), "pt-BR");
    });
  });

  return byRga;
}

function core_resolveChamadaLifecycleState_(events, refDate) {
  if (!events || !events.length) {
    return Object.freeze({
      hasHistory: false,
      isMemberOnDate: null,
      situacao: "",
      contaPresenca: null,
      contaFalta: null,
      motivoNaoAplicavel: ""
    });
  }

  const refTime = core_startOfDay_(refDate).getTime();
  let latest = null;
  let firstIngressAfterRef = null;

  events.forEach(function(event) {
    if (!event || !event.eventDate) return;
    const eventTime = core_startOfDay_(event.eventDate).getTime();
    const eventType = core_memberLifecycleNormalizeType_(event.eventType);

    if (eventTime <= refTime) {
      latest = event;
      return;
    }

    if (eventType === "INGRESSO" && !firstIngressAfterRef) {
      firstIngressAfterRef = event;
    }
  });

  if (!latest && firstIngressAfterRef) {
    return Object.freeze({
      hasHistory: true,
      isMemberOnDate: false,
      situacao: "NAO_INGRESSOU",
      contaPresenca: false,
      contaFalta: false,
      motivoNaoAplicavel: "INGRESSO_POSTERIOR_A_DATA"
    });
  }

  if (!latest) {
    return Object.freeze({
      hasHistory: true,
      isMemberOnDate: null,
      situacao: "",
      contaPresenca: null,
      contaFalta: null,
      motivoNaoAplicavel: ""
    });
  }

  const latestType = core_memberLifecycleNormalizeType_(latest.eventType);

  if (latestType === "INGRESSO" || latestType === "RETORNO") {
    return Object.freeze({
      hasHistory: true,
      isMemberOnDate: true,
      situacao: "ATIVO",
      contaPresenca: true,
      contaFalta: true,
      motivoNaoAplicavel: ""
    });
  }

  if (latestType === "SUSPENSAO") {
    return Object.freeze({
      hasHistory: true,
      isMemberOnDate: true,
      situacao: "SUSPENSO",
      contaPresenca: false,
      contaFalta: false,
      motivoNaoAplicavel: "SITUACAO_NAO_CONTABILIZA_CHAMADA"
    });
  }

  if (latestType.indexOf("DESLIGAMENTO") === 0) {
    return Object.freeze({
      hasHistory: true,
      isMemberOnDate: false,
      situacao: "DESLIGADO",
      contaPresenca: false,
      contaFalta: false,
      motivoNaoAplicavel: "VINCULO_ENCERRADO_ANTES_DA_DATA"
    });
  }

  return Object.freeze({
    hasHistory: true,
    isMemberOnDate: null,
    situacao: "",
    contaPresenca: null,
    contaFalta: null,
    motivoNaoAplicavel: ""
  });
}

function core_getChamadaLifecycleDisplayEvent_(events, refDate) {
  const refTime = core_startOfDay_(refDate).getTime();
  let selected = null;

  (events || []).forEach(function(event) {
    if (!event || !event.eventDate) return;
    if (core_startOfDay_(event.eventDate).getTime() > refTime) return;
    selected = event;
  });

  return selected || (events && events.length ? events[0] : null);
}

// Monta somente os campos permitidos para o Portal. Qualquer dado de contato,
// documento, observacao interna ou motivo sensivel fica fora deste contrato.
function core_buildChamadaMemberFromRow_(displayRow, rawRow, idx, refDate, lifecycleState) {
  const rga = core_getChamadaCell_(displayRow, idx.rga, "");
  const nomeExibicao = core_getChamadaCell_(displayRow, idx.name, "");
  const situacaoPlanilha = core_getChamadaCell_(displayRow, idx.situacaoGeral, "") ||
    core_getChamadaCell_(displayRow, idx.status, "") ||
    "ATIVO";
  const vinculo = core_getChamadaCell_(displayRow, idx.vinculo, "Membro");
  const dataIntegracao = core_parseChamadaRowDate_(rawRow, idx.dataIntegracao);
  const dataDesligamento = core_parseChamadaRowDate_(rawRow, idx.dataDesligamento);
  const ref = core_startOfDay_(refDate);

  if (dataIntegracao && core_startOfDay_(dataIntegracao).getTime() > ref.getTime()) {
    return null;
  }

  if (dataDesligamento && core_startOfDay_(dataDesligamento).getTime() < ref.getTime()) {
    return null;
  }

  if (lifecycleState && lifecycleState.isMemberOnDate === false) {
    return null;
  }

  const situacao = lifecycleState && lifecycleState.situacao
    ? lifecycleState.situacao
    : core_normalizeChamadaSituacao_(situacaoPlanilha) || "ATIVO";
  const counting = lifecycleState && lifecycleState.contaPresenca !== null
    ? lifecycleState
    : core_getChamadaCountingRuleFromSituation_(situacao);

  if (counting.motivoNaoAplicavel === "VINCULO_NAO_ATIVO_NA_DATA") {
    return null;
  }

  return Object.freeze({
    tipoParticipante: "MEMBRO",
    rga: String(rga || "").trim(),
    nomeExibicao: String(nomeExibicao || "").trim(),
    situacao: situacao,
    vinculo: String(vinculo || "Membro").trim(),
    aplicavelNaData: true,
    contaPresenca: counting.contaPresenca !== false,
    contaFalta: counting.contaFalta !== false,
    motivoNaoAplicavel: String(counting.motivoNaoAplicavel || "").trim()
  });
}

function core_buildChamadaMemberFromLifecycle_(events, refDate, lifecycleState) {
  const displayEvent = core_getChamadaLifecycleDisplayEvent_(events, refDate);
  if (!displayEvent || !displayEvent.rga || !displayEvent.memberName) return null;
  if (!lifecycleState || lifecycleState.isMemberOnDate !== true) return null;

  return Object.freeze({
    tipoParticipante: "MEMBRO",
    rga: String(displayEvent.rga || "").trim(),
    nomeExibicao: String(displayEvent.memberName || "").trim(),
    situacao: lifecycleState.situacao || "ATIVO",
    vinculo: "Membro",
    aplicavelNaData: true,
    contaPresenca: lifecycleState.contaPresenca !== false,
    contaFalta: lifecycleState.contaFalta !== false,
    motivoNaoAplicavel: String(lifecycleState.motivoNaoAplicavel || "").trim()
  });
}

function core_hasOnlySafeChamadaFields_(item) {
  const allowed = [
    "tipoParticipante",
    "rga",
    "nomeExibicao",
    "situacao",
    "vinculo",
    "aplicavelNaData",
    "contaPresenca",
    "contaFalta",
    "motivoNaoAplicavel"
  ].sort();

  return Object.keys(item || {}).sort().join(",") === allowed.join(",");
}

function core_listarMembrosParaChamadaInSheet_(sheet, dataAtividade, contexto, opts) {
  opts = opts || {};
  const refDate = core_parseChamadaReferenceDate_(dataAtividade);

  if (!refDate) {
    return core_buildChamadaError_(
      "DATA_ATIVIDADE_OBRIGATORIA",
      "Informe a data da atividade."
    );
  }

  if (!core_isChamadaAuthorizedContext_(contexto)) {
    return core_buildChamadaError_(
      "PERMISSAO_NEGADA",
      "Usuario sem permissao para listar membros para chamada."
    );
  }

  if (!sheet) {
    return core_buildChamadaError_(
      "BASE_MEMBROS_INDISPONIVEL",
      "Base de membros indisponivel para chamada."
    );
  }

  const observacoes = [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= CORE_MEMBERS_CFG.headerRow || lastCol < 1) {
    return Object.freeze({
      ok: true,
      data: Object.freeze([]),
      meta: Object.freeze({
        total: 0,
        dataReferencia: core_formatChamadaDateIso_(refDate),
        origemDados: "GEAPA_CORE",
        observacoes: Object.freeze(["MEMBERS_ATUAIS sem linhas de membros para chamada."])
      })
    });
  }

  const headers = sheet
    .getRange(CORE_MEMBERS_CFG.headerRow, 1, 1, lastCol)
    .getValues()[0]
    .map(function(header) {
      return String(header || "").trim();
    });
  const idx = core_getAttendanceMemberHeaderIndexMap_(headers);

  if (idx.name < 0 || idx.rga < 0) {
    return core_buildChamadaError_(
      "SCHEMA_MEMBROS_INVALIDO",
      "Base de membros sem cabecalhos obrigatorios para chamada."
    );
  }

  if (idx.dataIntegracao < 0) {
    observacoes.push(
      "Campo oficial de data de ingresso/integracao nao encontrado em MEMBERS_ATUAIS; quando nao houver historico de eventos, membros atuais podem ser considerados aplicaveis sem corte historico de entrada."
    );
  }

  if (idx.dataDesligamento < 0) {
    observacoes.push(
      "Campo oficial de data de desligamento/saida nao encontrado em MEMBERS_ATUAIS; desligamentos historicos dependem de MEMBER_EVENTOS_VINCULO."
    );
  }

  const startRow = CORE_MEMBERS_CFG.headerRow + 1;
  const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol);
  const values = range.getValues();
  const displayValues = typeof range.getDisplayValues === "function"
    ? range.getDisplayValues()
    : values;
  const lifecycleByRga = opts.lifecycleByRga || core_getChamadaLifecycleEventsByRga_(observacoes);
  const out = [];
  const seenRga = {};

  for (let i = 0; i < values.length; i++) {
    const rawRow = values[i];
    const displayRow = displayValues[i] || rawRow;
    const rga = core_getChamadaCell_(displayRow, idx.rga, "");
    const nome = core_getChamadaCell_(displayRow, idx.name, "");
    if (!rga || !nome) continue;

    const lifecycleKey = core_normalizeIdentityKey_(rga);
    seenRga[lifecycleKey] = true;
    const lifecycleState = core_resolveChamadaLifecycleState_(
      lifecycleByRga[lifecycleKey] || [],
      refDate
    );
    const item = core_buildChamadaMemberFromRow_(
      displayRow,
      rawRow,
      idx,
      refDate,
      lifecycleState
    );

    if (!item) continue;
    out.push(item);
  }

  let addedFromLifecycle = 0;
  Object.keys(lifecycleByRga).forEach(function(rgaKey) {
    if (seenRga[rgaKey]) return;

    const events = lifecycleByRga[rgaKey] || [];
    const lifecycleState = core_resolveChamadaLifecycleState_(events, refDate);
    const item = core_buildChamadaMemberFromLifecycle_(events, refDate, lifecycleState);
    if (!item) return;

    seenRga[rgaKey] = true;
    addedFromLifecycle++;
    out.push(item);
  });

  if (addedFromLifecycle > 0) {
    observacoes.push(
      "Alguns participantes foram reconstruidos por MEMBER_EVENTOS_VINCULO porque nao estavam em MEMBERS_ATUAIS na data da consulta; vinculo retorna como Membro por seguranca."
    );
  }

  out.sort(function(a, b) {
    return String(a.nomeExibicao || "").localeCompare(String(b.nomeExibicao || ""), "pt-BR");
  });

  return Object.freeze({
    ok: true,
    data: Object.freeze(out),
    meta: Object.freeze({
      total: out.length,
      dataReferencia: core_formatChamadaDateIso_(refDate),
      origemDados: "GEAPA_CORE",
      observacoes: Object.freeze(observacoes)
    })
  });
}

function core_listarMembrosParaChamada_(dataAtividade, contexto) {
  return core_listarMembrosParaChamadaInSheet_(
    core_getMembersSheet_(),
    dataAtividade,
    contexto || {}
  );
}

function core_runTesteListarMembrosParaChamada_() {
  const result = core_listarMembrosParaChamada_(core_formatChamadaDateIso_(new Date()), {
    perfil: "DIRETORIA"
  });

  if (!result || !result.ok) return result;

  const primeiraPessoa = result.data && result.data.length ? result.data[0] : "";
  const camposSeguros = !primeiraPessoa || core_hasOnlySafeChamadaFields_(primeiraPessoa);

  return Object.freeze({
    ok: true,
    total: result.meta ? result.meta.total : 0,
    primeiraPessoa: primeiraPessoa || "",
    camposSeguros: camposSeguros
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
