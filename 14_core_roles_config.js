/***************************************
 * 14_core_roles_config.js
 *
 * Leitura e indexação da tabela
 * CARGOS_INSTITUCIONAIS_CONFIG
 ***************************************/

const CORE_ROLES_CONFIG_KEY = 'CARGOS_INSTITUCIONAIS_CONFIG';
const CORE_ROLES_CONFIG_CACHE_KEY = 'GEAPA_CORE_ROLES_CONFIG_V1';
const CORE_ROLES_CONFIG_CACHE_TTL_SECONDS = 15 * 60;

/**
 * ------------------------------------------------------------
 * Abre a sheet da configuração de cargos institucionais.
 * ------------------------------------------------------------
 */
function core_getInstitutionalRolesConfigSheet_() {
  return core_getSheetByKey_(CORE_ROLES_CONFIG_KEY);
}

/**
 * ------------------------------------------------------------
 * Normaliza texto para comparações internas.
 * ------------------------------------------------------------
 */
function core_rolesNormalizeText_(value) {
  return String(value == null ? '' : value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

/**
 * ------------------------------------------------------------
 * Parseia SIM/NÃO.
 * ------------------------------------------------------------
 */
function core_rolesParseYesNo_(value) {
  const s = core_rolesNormalizeText_(value);
  if (!s) return false;
  return s === 'SIM';
}

/**
 * ------------------------------------------------------------
 * Parseia número de ordenação.
 * ------------------------------------------------------------
 */
function core_rolesParseDisplayOrder_(value, fallback) {
  if (value === '' || value == null) return fallback;
  const n = Number(value);
  return isNaN(n) ? fallback : n;
}

/**
 * ------------------------------------------------------------
 * Parseia lista CSV simples.
 * Ex.: "SECRETARIA,DIRETORIA"
 * ------------------------------------------------------------
 */
function core_rolesParseCsvList_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return [];

  return raw
    .split(',')
    .map(s => core_rolesNormalizeText_(s))
    .filter(Boolean);
}

/**
 * ------------------------------------------------------------
 * Parseia aliases/variações.
 * Ex.: "Vice Presidente, Vice-Presidente"
 * ------------------------------------------------------------
 */
function core_rolesParseAliases_(value) {
  return core_rolesParseCsvList_(value);
}

/**
 * ------------------------------------------------------------
 * Lê todas as linhas válidas da config e devolve objetos já
 * normalizados.
 * ------------------------------------------------------------
 *
 * Retorno:
 * [
 *   {
 *     publicName,
 *     publicNameNorm,
 *     group,
 *     groupNorm,
 *     roleKey,
 *     roleKeyNorm,
 *     isActive,
 *     displayOrder,
 *     receivesEmails,
 *     emailGroups,
 *     isUniqueRole,
 *     aliases,
 *     aliasesNorm
 *   }
 * ]
 */
function core_getInstitutionalRolesConfigRows_() {
  const cached = core_rolesConfigCacheGet_();
  if (cached) return cached;

  const sh = core_getInstitutionalRolesConfigSheet_();
  if (!sh) {
    throw new Error(
      `Sheet da config institucional não encontrada para KEY "${CORE_ROLES_CONFIG_KEY}".`
    );
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2 || lastCol < 1) {
    return [];
  }

  const headerMap = core_headerMap_(sh, 1);

  const colPublicName   = core_getCol_(headerMap, 'NOME_PUBLICO');
  const colGroup        = core_getCol_(headerMap, 'GRUPO_CARGO');
  const colRoleKey      = core_getCol_(headerMap, 'CARGO_KEY');
  const colAtivo        = core_getCol_(headerMap, 'ATIVO');
  const colDisplayOrder = core_getCol_(headerMap, 'DISPLAY_ORDEM');
  const colRecebeEmails = core_getCol_(headerMap, 'RECEBE_EMAILS');
  const colEmailsGrupo  = core_getCol_(headerMap, 'EMAILS_GRUPO');
  const colCargoUnico   = core_getCol_(headerMap, 'É_CARGO_UNICO');
  const colVariacao     = core_getCol_(headerMap, 'ESCRITA_VARIACAO');

  const required = [
    ['NOME_PUBLICO', colPublicName],
    ['GRUPO_CARGO', colGroup],
    ['CARGO_KEY', colRoleKey],
    ['ATIVO', colAtivo],
    ['DISPLAY_ORDEM', colDisplayOrder],
    ['RECEBE_EMAILS', colRecebeEmails],
    ['EMAILS_GRUPO', colEmailsGrupo],
    ['É_CARGO_UNICO', colCargoUnico],
    ['ESCRITA_VARIACAO', colVariacao]
  ];

  const missing = required
    .filter(item => !item[1])
    .map(item => item[0]);

  if (missing.length) {
    throw new Error(
      'CARGOS_INSTITUCIONAIS_CONFIG inválida. Cabeçalhos ausentes: ' + missing.join(', ')
    );
  }

  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const rows = values
    .map((row, idx) => {
      const lineNo = idx + 2;

      const publicName = String(row[colPublicName - 1] || '').trim();
      const group = String(row[colGroup - 1] || '').trim();
      const roleKey = String(row[colRoleKey - 1] || '').trim();

      if (!publicName || !group || !roleKey) {
        return null;
      }

      const aliases = core_rolesParseAliases_(row[colVariacao - 1]);
      const publicNameNorm = core_rolesNormalizeText_(publicName);
      const groupNorm = core_rolesNormalizeText_(group);
      const roleKeyNorm = core_rolesNormalizeText_(roleKey);

      return Object.freeze({
        lineNo: lineNo,
        publicName: publicName,
        publicNameNorm: publicNameNorm,
        group: group,
        groupNorm: groupNorm,
        roleKey: roleKey,
        roleKeyNorm: roleKeyNorm,
        isActive: core_rolesParseYesNo_(row[colAtivo - 1]),
        displayOrder: core_rolesParseDisplayOrder_(row[colDisplayOrder - 1], 999999),
        receivesEmails: core_rolesParseYesNo_(row[colRecebeEmails - 1]),
        emailGroups: Object.freeze(core_rolesParseCsvList_(row[colEmailsGrupo - 1])),
        isUniqueRole: core_rolesParseYesNo_(row[colCargoUnico - 1]),
        aliases: Object.freeze(aliases),
        aliasesNorm: Object.freeze(aliases.map(a => core_rolesNormalizeText_(a)))
      });
    })
    .filter(Boolean);

  const frozen = Object.freeze(rows);
  core_rolesConfigCacheSet_(frozen);
  return frozen;
}

/**
 * ------------------------------------------------------------
 * Monta mapa/index da config para buscas rápidas.
 * ------------------------------------------------------------
 */
function core_getInstitutionalRolesConfigMap_() {
  const rows = core_getInstitutionalRolesConfigRows_();

  const byKey = {};
  const byPublicName = {};
  const byAnyName = {};
  const byEmailGroup = {};

  rows.forEach((role) => {
    if (!role.isActive) return;

    byKey[role.roleKeyNorm] = role;
    byPublicName[role.publicNameNorm] = role;

    // busca por qualquer escrita conhecida
    byAnyName[role.roleKeyNorm] = role;
    byAnyName[role.publicNameNorm] = role;
    role.aliasesNorm.forEach(aliasNorm => {
      byAnyName[aliasNorm] = role;
    });

    role.emailGroups.forEach(groupNameNorm => {
      if (!byEmailGroup[groupNameNorm]) byEmailGroup[groupNameNorm] = [];
      byEmailGroup[groupNameNorm].push(role);
    });
  });

  Object.keys(byEmailGroup).forEach((k) => {
    byEmailGroup[k] = Object.freeze(
      byEmailGroup[k].slice().sort((a, b) => a.displayOrder - b.displayOrder)
    );
  });

  return Object.freeze({
    rows: rows,
    byKey: Object.freeze(byKey),
    byPublicName: Object.freeze(byPublicName),
    byAnyName: Object.freeze(byAnyName),
    byEmailGroup: Object.freeze(byEmailGroup)
  });
}

/**
 * ------------------------------------------------------------
 * Busca cargo por CARGO_KEY.
 * ------------------------------------------------------------
 */
function core_findInstitutionalRoleByKey_(cargoKey) {
  core_assertRequired_(cargoKey, 'cargoKey');
  const map = core_getInstitutionalRolesConfigMap_();
  return map.byKey[core_rolesNormalizeText_(cargoKey)] || null;
}

/**
 * ------------------------------------------------------------
 * Busca cargo por NOME_PUBLICO exato normalizado.
 * ------------------------------------------------------------
 */
function core_findInstitutionalRoleByPublicName_(publicName) {
  core_assertRequired_(publicName, 'publicName');
  const map = core_getInstitutionalRolesConfigMap_();
  return map.byPublicName[core_rolesNormalizeText_(publicName)] || null;
}

/**
 * ------------------------------------------------------------
 * Busca cargo por qualquer nome conhecido:
 * - CARGO_KEY
 * - NOME_PUBLICO
 * - ESCRITA_VARIACAO
 * ------------------------------------------------------------
 */
function core_findInstitutionalRoleByAnyName_(text) {
  core_assertRequired_(text, 'text');
  const map = core_getInstitutionalRolesConfigMap_();
  return map.byAnyName[core_rolesNormalizeText_(text)] || null;
}

/**
 * ------------------------------------------------------------
 * Retorna cargos ativos de um grupo de e-mail.
 * Ex.: SECRETARIA, COMUNICACAO, EVENTOS...
 * ------------------------------------------------------------
 */
function core_getInstitutionalRolesByEmailGroup_(groupName) {
  core_assertRequired_(groupName, 'groupName');
  const map = core_getInstitutionalRolesConfigMap_();
  return map.byEmailGroup[core_rolesNormalizeText_(groupName)] || Object.freeze([]);
}

/**
 * ------------------------------------------------------------
 * Retorna todos os cargos ativos.
 * ------------------------------------------------------------
 */
function core_getInstitutionalRolesActive_() {
  return Object.freeze(
    core_getInstitutionalRolesConfigRows_()
      .filter(r => r.isActive)
      .slice()
      .sort((a, b) => a.displayOrder - b.displayOrder)
  );
}

/**
 * ------------------------------------------------------------
 * Debug / healthcheck da configuração.
 * ------------------------------------------------------------
 */
function core_debugInstitutionalRolesConfig_() {
  const all = core_getInstitutionalRolesConfigRows_();
  const active = core_getInstitutionalRolesActive_();
  const secretariat = core_getInstitutionalRolesByEmailGroup_('SECRETARIA');
  const communication = core_getInstitutionalRolesByEmailGroup_('COMUNICACAO');
  const events = core_getInstitutionalRolesByEmailGroup_('EVENTOS');

  return {
    totalRows: all.length,
    activeRows: active.length,
    receivesEmails: active.filter(r => r.receivesEmails).length,
    groups: {
      SECRETARIA: secretariat.map(r => r.publicName),
      COMUNICACAO: communication.map(r => r.publicName),
      EVENTOS: events.map(r => r.publicName)
    }
  };
}

/**
 * ------------------------------------------------------------
 * Cache GET
 * ------------------------------------------------------------
 */
function core_rolesConfigCacheGet_() {
  const cache = CacheService.getScriptCache();
  const s = cache.get(CORE_ROLES_CONFIG_CACHE_KEY);
  if (!s) return null;

  try {
    const arr = JSON.parse(s);
    return Object.freeze(arr.map(item => Object.freeze({
      lineNo: item.lineNo,
      publicName: item.publicName,
      publicNameNorm: item.publicNameNorm,
      group: item.group,
      groupNorm: item.groupNorm,
      roleKey: item.roleKey,
      roleKeyNorm: item.roleKeyNorm,
      isActive: item.isActive,
      displayOrder: item.displayOrder,
      receivesEmails: item.receivesEmails,
      emailGroups: Object.freeze(item.emailGroups || []),
      isUniqueRole: item.isUniqueRole,
      aliases: Object.freeze(item.aliases || []),
      aliasesNorm: Object.freeze(item.aliasesNorm || [])
    })));
  } catch (e) {
    return null;
  }
}

/**
 * ------------------------------------------------------------
 * Cache SET
 * ------------------------------------------------------------
 */
function core_rolesConfigCacheSet_(rows) {
  CacheService
    .getScriptCache()
    .put(
      CORE_ROLES_CONFIG_CACHE_KEY,
      JSON.stringify(rows),
      CORE_ROLES_CONFIG_CACHE_TTL_SECONDS
    );
}

/**
 * ------------------------------------------------------------
 * Cache CLEAR
 * ------------------------------------------------------------
 */
function core_rolesConfigCacheClear_() {
  CacheService.getScriptCache().remove(CORE_ROLES_CONFIG_CACHE_KEY);
}