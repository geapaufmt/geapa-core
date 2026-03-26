/*******************************
 *****  Registry Dinâmico  *****
 * Lê uma planilha "mãe" (GEAPA_REGISTRY) e monta:
 * { KEY: { id, sheet, ...meta } }
 *******************************/

// ✅ 1 único ID fixo no core (planilha-mãe)
const CORE_REGISTRY_SPREADSHEET_ID = '1KQ_-GcFuvLA-jrPVQsfE3_TuaR8AM0yM4X41AzEGXGI';
// Nome fixo da aba dentro dessa planilha
const CORE_REGISTRY_SHEET_NAME = 'Registry';

// Cache
const CORE_REGISTRY_CACHE_KEY = 'GEAPA_CORE_REGISTRY_V2_RAW';
const CORE_REGISTRY_CACHE_TTL_SECONDS = 15 * 60; // 15 min

// Ambiente default se não houver Script Property
const CORE_DEFAULT_ENV = 'PROD';

/**
 * ------------------------------------------------------------
 * Retorna o ambiente atual do sistema.
 * ------------------------------------------------------------
 *
 * Fonte preferencial:
 * - Script Properties: GEAPA_ENV = DEV | PROD
 *
 * Fallback:
 * - CORE_DEFAULT_ENV
 */
function core_getCurrentEnv_() {
  const props = PropertiesService.getScriptProperties();
  const raw = String(props.getProperty('GEAPA_ENV') || CORE_DEFAULT_ENV).trim().toUpperCase();

  if (raw === 'DEV' || raw === 'PROD') return raw;

  throw new Error(
    `GEAPA_ENV inválido: "${raw}". Valores aceitos: DEV ou PROD.`
  );
}

/**
 * ------------------------------------------------------------
 * True se a entrada puder ser usada no ambiente atual.
 * ------------------------------------------------------------
 */
function core_registryMatchesEnv_(entryEnv, currentEnv) {
  if (!entryEnv || entryEnv === 'ALL') return true;
  return entryEnv === currentEnv;
}

/**
 * ------------------------------------------------------------
 * Lê o registry bruto, preservando metadados.
 * ------------------------------------------------------------
 *
 * Retorna:
 * {
 *   KEY: {
 *     key,
 *     id,
 *     sheet,
 *     displayName,
 *     ativo,
 *     ambiente,
 *     type,
 *     notas,
 *     lineNo
 *   }
 * }
 */
function core_getRegistryRaw_() {
  const runId = core_runId_();
  const startedAt = new Date();

  const cached = core_registryCacheGet_();
  if (cached) {
    core_logInfo_(runId, 'Registry RAW (cache): OK', { keys: Object.keys(cached).length });
    core_logSummarize_(runId, 'core_getRegistryRaw_', startedAt, { source: 'cache' });
    return cached;
  }

  core_logInfo_(runId, 'Registry RAW (cache): MISS, lendo planilha-mãe');

  const ss = core_openSpreadsheetById_(CORE_REGISTRY_SPREADSHEET_ID);
  const sh = ss.getSheetByName(CORE_REGISTRY_SHEET_NAME);
  if (!sh) {
    throw new Error(
      `Aba "${CORE_REGISTRY_SHEET_NAME}" não encontrada na planilha Registry (${CORE_REGISTRY_SPREADSHEET_ID}).`
    );
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    throw new Error('Planilha Registry vazia: precisa de cabeçalho e ao menos 1 linha de dados.');
  }

  const headerMap = core_headerMap_(sh, 1);

  const colKey       = core_getCol_(headerMap, 'KEY');
  const colId        = core_getCol_(headerMap, 'SPREADSHEET_ID');
  const colSheet     = core_getCol_(headerMap, 'SHEET_NAME');
  const colAtivo     = core_getCol_(headerMap, 'ATIVO');
  const colAmbiente  = core_getCol_(headerMap, 'AMBIENTE');
  const colDisplay   = core_getCol_(headerMap, 'DISPLAY_NAME');
  const colType      = core_getCol_(headerMap, 'TYPE');
  const colNotas     = core_getCol_(headerMap, 'NOTAS');

  if (!colKey || !colId || !colSheet || !colAtivo || !colAmbiente) {
    throw new Error(
      'Registry inválido. Cabeçalhos obrigatórios: KEY, SPREADSHEET_ID, SHEET_NAME, ATIVO, AMBIENTE.'
    );
  }

  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const out = {};
  let ignorados = 0;
  let invalidos = 0;

  values.forEach((row, idx) => {
    const lineNo = idx + 2;
    const key = String(row[colKey - 1] || '').trim().toUpperCase();
    const id = String(row[colId - 1] || '').trim();
    const sheet = String(row[colSheet - 1] || '').trim();

    if (!key) {
      ignorados++;
      return;
    }

    if (!id || !sheet) {
      invalidos++;
      core_logWarn_(runId, 'Registry linha inválida (faltando id/sheet)', { lineNo, key, id, sheet });
      return;
    }

    const ativo = core_registryParseRequiredAtivo_(
      row[colAtivo - 1],
      key,
      lineNo
    );

    const ambiente = core_registryParseRequiredAmbiente_(
      row[colAmbiente - 1],
      key,
      lineNo
    );

    const displayName = colDisplay ? String(row[colDisplay - 1] || '').trim() : '';
    const type = colType ? String(row[colType - 1] || '').trim().toUpperCase() : '';
    const notas = colNotas ? String(row[colNotas - 1] || '').trim() : '';

    if (!out[key]) out[key] = {};

    if (out[key][ambiente]) {
      invalidos++;
      core_logWarn_(runId, 'Registry duplicado no mesmo ambiente (mantendo o primeiro)', {
        lineNo,
        key,
        ambiente
      });
      return;
    }

    out[key][ambiente] = Object.freeze({
      key,
      id,
      sheet,
      displayName,
      ativo,
      ambiente,
      type,
      notas,
      lineNo
    });
  });

  const frozen = {};
  Object.keys(out).forEach((key) => {
    frozen[key] = Object.freeze(out[key]);
  });

  const fullyFrozen = Object.freeze(frozen);
  core_registryCacheSet_(fullyFrozen);

  core_logInfo_(runId, 'Registry RAW (planilha): OK', {
    keys: Object.keys(fullyFrozen).length,
    ignorados,
    invalidos
  });

  core_logSummarize_(runId, 'core_getRegistryRaw_', startedAt, { source: 'sheet' });

  return fullyFrozen;
}

/**
 * ------------------------------------------------------------
 * Retorna o registry filtrado para o ambiente atual.
 * ------------------------------------------------------------
 *
 * Mantém compatibilidade com quem espera um mapa simples:
 * { KEY: { id, sheet } }
 */
function core_getRegistry_() {
  const currentEnv = core_getCurrentEnv_();
  const raw = core_getRegistryRaw_();

  const out = {};
  Object.keys(raw).forEach((k) => {
    const envMap = raw[k];
    const entry = envMap[currentEnv];
    if (!entry) return;
    if (!entry.ativo) return;

    out[k] = Object.freeze({
      id: entry.id,
      sheet: entry.sheet
    });
  });

  return Object.freeze(out);
}

/**
 * ------------------------------------------------------------
 * Busca registry bruto no cache.
 * ------------------------------------------------------------
 */
function core_registryCacheGet_() {
  const cache = CacheService.getScriptCache();
  const s = cache.get(CORE_REGISTRY_CACHE_KEY);
  if (!s) return null;

  try {
    const obj = JSON.parse(s);
    const out = {};

    Object.keys(obj).forEach((k) => {
      out[k] = {};

      Object.keys(obj[k]).forEach((ambiente) => {
        out[k][ambiente] = Object.freeze({
          key: obj[k][ambiente].key,
          id: obj[k][ambiente].id,
          sheet: obj[k][ambiente].sheet,
          displayName: obj[k][ambiente].displayName,
          ativo: obj[k][ambiente].ativo,
          ambiente: obj[k][ambiente].ambiente,
          type: obj[k][ambiente].type,
          notas: obj[k][ambiente].notas,
          lineNo: obj[k][ambiente].lineNo
        });
      });

      out[k] = Object.freeze(out[k]);
    });

    return Object.freeze(out);
  } catch (e) {
    return null;
  }
}

/**
 * ------------------------------------------------------------
 * Salva registry bruto no cache.
 * ------------------------------------------------------------
 */
function core_registryCacheSet_(registryObj) {
  const cache = CacheService.getScriptCache();
  cache.put(CORE_REGISTRY_CACHE_KEY, JSON.stringify(registryObj), CORE_REGISTRY_CACHE_TTL_SECONDS);
}

/**
 * ------------------------------------------------------------
 * Limpa cache do registry.
 * ------------------------------------------------------------
 */
function core_registryCacheClear_() {
  CacheService.getScriptCache().remove(CORE_REGISTRY_CACHE_KEY);
}

/**
 * ------------------------------------------------------------
 * Retorna metadados completos de uma KEY.
 * ------------------------------------------------------------
 */
function core_getRegistryMetaByKey_(key) {
  core_assertRequired_(key, 'Registry KEY');

  const k = String(key).trim().toUpperCase();
  const raw = core_getRegistryRaw_();
  const envMap = raw[k];

  if (!envMap) {
    const keys = Object.keys(raw).sort();
    throw new Error(
      `Registry KEY não encontrada: "${k}". ` +
      `Cadastre na planilha Registry ou corrija a KEY. ` +
      `Disponíveis: ${keys.join(', ')}`
    );
  }

  const currentEnv = core_getCurrentEnv_();
  const entry = envMap[currentEnv];

  if (!entry) {
    const ambientesDisponiveis = Object.keys(envMap).sort();
    throw new Error(
      `Registry KEY "${k}" não possui cadastro para o ambiente atual "${currentEnv}". ` +
      `Ambientes disponíveis para essa KEY: ${ambientesDisponiveis.join(', ')}`
    );
  }

  return entry;
}
/**
 * ------------------------------------------------------------
 * Retorna a referência do Registry por KEY, validando:
 * - existência
 * - ativo/inativo
 * - compatibilidade com ambiente atual
 * ------------------------------------------------------------
 */
function core_getRegistryRefByKey_(key) {
  const entry = core_getRegistryMetaByKey_(key);
  const currentEnv = core_getCurrentEnv_();

  if (!entry.ativo) {
    throw new Error(
      `Registry KEY "${entry.key}" está inativa (linha ${entry.lineNo}).`
    );
  }

  if (!core_registryMatchesEnv_(entry.ambiente, currentEnv)) {
    throw new Error(
      `Registry KEY "${entry.key}" não está liberada para o ambiente atual. ` +
      `Ambiente da entry: "${entry.ambiente || 'vazio'}". ` +
      `Ambiente atual: "${currentEnv}".`
    );
  }

  return Object.freeze({ id: entry.id, sheet: entry.sheet });
}

/**
 * ------------------------------------------------------------
 * Abre e retorna o Sheet diretamente a partir de uma KEY.
 * ------------------------------------------------------------
 */
function core_getSheetByKey_(key) {
  const ref = core_getRegistryRefByKey_(key);
  return core_getSheetById_(ref.id, ref.sheet);
}

/**
 * ------------------------------------------------------------
 * Valida e interpreta a coluna ATIVO (obrigatória).
 * ------------------------------------------------------------
 *
 * Regras:
 * - SIM  => true
 * - NÃO  => false
 * - qualquer outro valor => ERRO
 */
function core_registryParseRequiredAtivo_(value, key, lineNo) {
  const s = String(value || '').trim().toUpperCase();

  if (s === 'SIM') return true;
  if (s === 'NÃO' || s === 'NAO') return false;

  throw new Error(
    `Registry inválido na linha ${lineNo} (KEY="${key}"):\n` +
    `A coluna ATIVO deve ser "SIM" ou "NÃO".\n` +
    `Valor recebido: "${s || '(vazio)'}"`
  );
}

/**
 * ------------------------------------------------------------
 * Valida e interpreta a coluna AMBIENTE (opcional).
 * ------------------------------------------------------------
 *
 * Regras:
 * - DEV  => 'DEV'
 * - PROD => 'PROD'
 * - qualquer outro valor => ERRO
 */

function core_registryParseRequiredAmbiente_(value, key, lineNo) {
  const s = String(value || '').trim().toUpperCase();

  if (s === 'DEV') return 'DEV';
  if (s === 'PROD') return 'PROD';

  throw new Error(
    `Registry inválido na linha ${lineNo} (KEY="${key}"):\n` +
    `A coluna AMBIENTE deve ser "DEV" ou "PROD".\n` +
    `Valor recebido: "${s || '(vazio)'}"`
  );
}
