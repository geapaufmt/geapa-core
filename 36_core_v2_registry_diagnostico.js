/**
 * ============================================================
 * 36_core_v2_registry_diagnostico.js
 * ============================================================
 *
 * Diagnostico somente leitura das keys V2 no Registry bruto.
 *
 * Esta rotina nao escreve em planilhas, nao cria triggers e nao abre as bases
 * de destino. O objetivo e distinguir "key ausente" de "key existe em outro
 * ambiente" quando o Registry filtrado por GEAPA_ENV oculta linhas DEV.
 */

var CORE_V2_REGISTRY_DIAGNOSTICO_KEYS = Object.freeze([
  'PESSOAS_V2_BASE',
  'PESSOAS_V2_IDENTIFICADORES',
  'PESSOAS_V2_MEMBROS_DETALHES',
  'PESSOAS_V2_VINCULOS_GEAPA',
  'PESSOAS_V2_MEMBROS_EVENTOS_VINCULO',
  'PESSOAS_V2_RESUMO_OPERACIONAL',
  'VIGENCIAS_V2_SEMESTRES',
  'VIGENCIAS_V2_PERIODOS',
  'VIGENCIAS_V2_DIRETORIAS',
  'VIGENCIAS_V2_SEMESTRES_DIRETORIA',
  'VIGENCIAS_V2_CARGOS_CONFIG',
  'VIGENCIAS_V2_FUNCOES',
  'VIGENCIAS_V2_RESUMO_ATUAL'
]);

function core_v2RegistryDiagnosticoOptions_(options) {
  options = options || {};
  var ambiente = core_v2RegistryDiagnosticoNormalize_(options.ambiente || 'DEV');
  return {
    ambiente: ambiente,
    ambienteValido: ambiente === 'DEV' || ambiente === 'PROD',
    incluirInativas: options.incluirInativas === true
  };
}

function core_v2RegistryDiagnosticoNormalize_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  }).replace(/\s+/g, '_');
}

function core_v2RegistryDiagnosticoMaskId_(value) {
  var id = String(value || '').trim();
  if (!id) return '';
  if (id.length <= 10) return '[id_mascarado]';
  return id.slice(0, 6) + '...' + id.slice(-4);
}

function core_v2RegistryDiagnosticoKey_(key, raw, opts) {
  var wanted = core_v2RegistryDiagnosticoNormalize_(key);
  var envMap = raw[wanted];
  if (!envMap) {
    return {
      key: wanted,
      encontrada: false,
      spreadsheetId: '',
      sheetName: '',
      ambiente: opts.ambiente,
      ativo: false,
      erro: 'KEY_NAO_ENCONTRADA'
    };
  }

  var entry = envMap[opts.ambiente];
  if (!entry) {
    return {
      key: wanted,
      encontrada: false,
      spreadsheetId: '',
      sheetName: '',
      ambiente: opts.ambiente,
      ativo: false,
      erro: 'KEY_SEM_AMBIENTE_' + opts.ambiente + '; ambientesDisponiveis=' + Object.keys(envMap).sort().join(',')
    };
  }

  return {
    key: wanted,
    encontrada: true,
    spreadsheetId: core_v2RegistryDiagnosticoMaskId_(entry.id),
    sheetName: entry.sheet || '',
    ambiente: entry.ambiente || opts.ambiente,
    ativo: entry.ativo === true,
    erro: entry.ativo === true || opts.incluirInativas
      ? ''
      : 'KEY_INATIVA'
  };
}

function core_v2ResolverRegistryV2_(options) {
  var opts = core_v2RegistryDiagnosticoOptions_(options || {});
  var result = {
    ok: true,
    modulo: 'CORE',
    fluxo: 'DIAGNOSTICO_REGISTRY_V2',
    readOnly: true,
    ambiente: opts.ambiente,
    ambienteValido: opts.ambienteValido,
    totalKeys: CORE_V2_REGISTRY_DIAGNOSTICO_KEYS.length,
    totalEncontradas: 0,
    totalAusentes: 0,
    totalInativas: 0,
    keys: [],
    erros: [],
    avisos: []
  };

  if (!opts.ambienteValido) {
    result.ok = false;
    result.erros.push('Ambiente invalido: ' + (opts.ambiente || '(vazio)') + '. Use DEV ou PROD.');
    return result;
  }

  try {
    var raw = core_getRegistryRaw_();
    result.keys = CORE_V2_REGISTRY_DIAGNOSTICO_KEYS.map(function(key) {
      return core_v2RegistryDiagnosticoKey_(key, raw, opts);
    });
  } catch (err) {
    result.ok = false;
    result.erros.push(core_v2RotinasSanitizeError_(err));
    return result;
  }

  result.totalEncontradas = result.keys.filter(function(item) { return item.encontrada; }).length;
  result.totalAusentes = result.keys.filter(function(item) { return !item.encontrada; }).length;
  result.totalInativas = result.keys.filter(function(item) { return item.encontrada && item.ativo !== true; }).length;

  if (result.totalAusentes) {
    result.ok = false;
    result.avisos.push('Existem keys V2 ausentes no ambiente ' + opts.ambiente + '.');
  }
  if (result.totalInativas) {
    result.ok = false;
    result.avisos.push('Existem keys V2 inativas no ambiente ' + opts.ambiente + '.');
  }

  return result;
}

function coreV2_runTesteResolverRegistryV2_() {
  return core_v2ResolverRegistryV2_({
    ambiente: 'DEV'
  });
}
