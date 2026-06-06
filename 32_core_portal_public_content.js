/***************************************
 * 32_core_portal_public_content.js
 *
 * Setup estrutural da planilha PORTAL_CONTEUDO_PUBLICO.
 *
 * Esta camada e administrativa, idempotente e nao destrutiva:
 * - cria abas editoriais ausentes;
 * - garante cabecalhos obrigatorios;
 * - adiciona colunas faltantes ao final;
 * - preserva dados e colunas extras;
 * - nao registra Registry automaticamente nesta etapa.
 ***************************************/

const CORE_PORTAL_PUBLIC_CONTENT_CFG = Object.freeze({
  spreadsheetName: 'PORTAL_CONTEUDO_PUBLICO',
  headerRow: 1,
  lockKey: 'PORTAL_CONTEUDO_PUBLICO_SETUP',
  registryKeys: Object.freeze({
    PUBLIC_HOME: 'PORTAL_PUBLIC_HOME',
    PUBLIC_SOBRE: 'PORTAL_PUBLIC_SOBRE',
    PUBLIC_HISTORIA: 'PORTAL_PUBLIC_HISTORIA',
    PUBLIC_PARCEIROS: 'PORTAL_PUBLIC_PARCEIROS',
    PUBLIC_DOCUMENTOS: 'PORTAL_PUBLIC_DOCUMENTOS',
    PUBLIC_CONFIG: 'PORTAL_PUBLIC_CONFIG',
    PUBLIC_MIDIAS: 'PORTAL_PUBLIC_MIDIAS',
    PUBLIC_DIRETORIA_COMPLEMENTOS: 'PORTAL_PUBLIC_DIRETORIA_COMPLEMENTOS',
    PUBLIC_LOG_PUBLICACAO: 'PORTAL_PUBLIC_LOG_PUBLICACAO'
  }),
  definitions: Object.freeze([
    Object.freeze({
      sheetName: 'PUBLIC_HOME',
      key: 'PORTAL_PUBLIC_HOME',
      headers: Object.freeze([
        'ID_BLOCO',
        'TIPO_BLOCO',
        'TITULO',
        'SUBTITULO',
        'TEXTO',
        'IMAGEM_URL',
        'BOTAO_TEXTO',
        'BOTAO_URL',
        'ORDEM',
        'STATUS_PUBLICACAO',
        'PUBLICAR',
        'ATIVO',
        'ATUALIZADO_EM'
      ])
    }),
    Object.freeze({
      sheetName: 'PUBLIC_SOBRE',
      key: 'PORTAL_PUBLIC_SOBRE',
      headers: Object.freeze([
        'ID_BLOCO',
        'TITULO',
        'TEXTO',
        'IMAGEM_URL',
        'ORDEM',
        'STATUS_PUBLICACAO',
        'PUBLICAR',
        'ATIVO',
        'ATUALIZADO_EM'
      ])
    }),
    Object.freeze({
      sheetName: 'PUBLIC_HISTORIA',
      key: 'PORTAL_PUBLIC_HISTORIA',
      headers: Object.freeze([
        'ID_MARCO',
        'TIPO',
        'ANO',
        'DATA',
        'TITULO',
        'TEXTO',
        'IMAGEM_URL',
        'ORDEM',
        'STATUS_PUBLICACAO',
        'PUBLICAR',
        'ATIVO',
        'ATUALIZADO_EM'
      ])
    }),
    Object.freeze({
      sheetName: 'PUBLIC_PARCEIROS',
      key: 'PORTAL_PUBLIC_PARCEIROS',
      headers: Object.freeze([
        'ID_PARCEIRO',
        'NOME',
        'TIPO_PARCEIRO',
        'DESCRICAO',
        'LOGO_URL',
        'SITE_URL',
        'INSTAGRAM_URL',
        'ORDEM',
        'STATUS_PUBLICACAO',
        'PUBLICAR',
        'ATIVO',
        'ATUALIZADO_EM'
      ])
    }),
    Object.freeze({
      sheetName: 'PUBLIC_DOCUMENTOS',
      key: 'PORTAL_PUBLIC_DOCUMENTOS',
      headers: Object.freeze([
        'ID_DOCUMENTO',
        'TITULO',
        'TIPO_DOCUMENTO',
        'VERSAO',
        'DATA_PUBLICACAO',
        'DESCRICAO',
        'URL_DOCUMENTO',
        'ORDEM',
        'STATUS_PUBLICACAO',
        'PUBLICAR',
        'ATIVO',
        'ATUALIZADO_EM'
      ])
    }),
    Object.freeze({
      sheetName: 'PUBLIC_CONFIG',
      key: 'PORTAL_PUBLIC_CONFIG',
      headers: Object.freeze([
        'KEY',
        'VALOR',
        'TIPO',
        'GRUPO',
        'DESCRICAO',
        'ATIVO'
      ])
    }),
    Object.freeze({
      sheetName: 'PUBLIC_MIDIAS',
      key: 'PORTAL_PUBLIC_MIDIAS',
      headers: Object.freeze([
        'ID_MIDIA',
        'NOME',
        'TIPO',
        'URL',
        'DESCRICAO',
        'CATEGORIA',
        'STATUS_PUBLICACAO',
        'PUBLICAR',
        'ATIVO',
        'ATUALIZADO_EM'
      ])
    }),
    Object.freeze({
      sheetName: 'PUBLIC_DIRETORIA_COMPLEMENTOS',
      key: 'PORTAL_PUBLIC_DIRETORIA_COMPLEMENTOS',
      headers: Object.freeze([
        'ID_PESSOA',
        'ID_DIRETORIA',
        'FOTO_URL',
        'DESCRICAO_PUBLICA',
        'LINK_LATTES',
        'LINK_INSTAGRAM_PUBLICO',
        'ORDEM_PUBLICA',
        'STATUS_PUBLICACAO',
        'PUBLICAR',
        'ATIVO',
        'ATUALIZADO_EM'
      ])
    }),
    Object.freeze({
      sheetName: 'PUBLIC_LOG_PUBLICACAO',
      key: 'PORTAL_PUBLIC_LOG_PUBLICACAO',
      headers: Object.freeze([
        'ID_LOG',
        'DATA_HORA',
        'TIPO_CONTEUDO',
        'ACAO',
        'EXECUTADO_POR',
        'STATUS',
        'OBSERVACAO'
      ])
    })
  ])
});

const CORE_PORTAL_PUBLIC_CONTENT_READ_CFG = Object.freeze({
  cacheTtlSeconds: 5 * 60,
  cachePrefix: 'GEAPA_CORE_PORTAL_PUBLIC_CONTENT_V1_',
  blockedPublicKeys: Object.freeze(['PORTAL_PUBLIC_LOG_PUBLICACAO']),
  pageSlugMap: Object.freeze({
    home: 'PORTAL_PUBLIC_HOME',
    sobre: 'PORTAL_PUBLIC_SOBRE',
    historia: 'PORTAL_PUBLIC_HISTORIA',
    parceiros: 'PORTAL_PUBLIC_PARCEIROS',
    documentos: 'PORTAL_PUBLIC_DOCUMENTOS',
    config: 'PORTAL_PUBLIC_CONFIG',
    midias: 'PORTAL_PUBLIC_MIDIAS',
    diretoriacomplementos: 'PORTAL_PUBLIC_DIRETORIA_COMPLEMENTOS'
  }),
  publicFields: Object.freeze({
    PORTAL_PUBLIC_HOME: Object.freeze([
      Object.freeze({ from: 'ID_BLOCO', to: 'idBloco', type: 'text' }),
      Object.freeze({ from: 'TIPO_BLOCO', to: 'tipoBloco', type: 'token' }),
      Object.freeze({ from: 'TITULO', to: 'titulo', type: 'text' }),
      Object.freeze({ from: 'SUBTITULO', to: 'subtitulo', type: 'text' }),
      Object.freeze({ from: 'TEXTO', to: 'texto', type: 'text' }),
      Object.freeze({ from: 'IMAGEM_URL', to: 'imagemUrl', type: 'url' }),
      Object.freeze({ from: 'BOTAO_TEXTO', to: 'botaoTexto', type: 'text' }),
      Object.freeze({ from: 'BOTAO_URL', to: 'botaoUrl', type: 'url' }),
      Object.freeze({ from: 'ORDEM', to: 'ordem', type: 'number' }),
      Object.freeze({ from: 'ATUALIZADO_EM', to: 'atualizadoEm', type: 'date' })
    ]),
    PORTAL_PUBLIC_SOBRE: Object.freeze([
      Object.freeze({ from: 'ID_BLOCO', to: 'idBloco', type: 'text' }),
      Object.freeze({ from: 'TITULO', to: 'titulo', type: 'text' }),
      Object.freeze({ from: 'TEXTO', to: 'texto', type: 'text' }),
      Object.freeze({ from: 'IMAGEM_URL', to: 'imagemUrl', type: 'url' }),
      Object.freeze({ from: 'ORDEM', to: 'ordem', type: 'number' }),
      Object.freeze({ from: 'ATUALIZADO_EM', to: 'atualizadoEm', type: 'date' })
    ]),
    PORTAL_PUBLIC_HISTORIA: Object.freeze([
      Object.freeze({ from: 'ID_MARCO', to: 'idMarco', type: 'text' }),
      Object.freeze({ from: 'TIPO', to: 'tipo', type: 'token' }),
      Object.freeze({ from: 'ANO', to: 'ano', type: 'text' }),
      Object.freeze({ from: 'DATA', to: 'data', type: 'date' }),
      Object.freeze({ from: 'TITULO', to: 'titulo', type: 'text' }),
      Object.freeze({ from: 'TEXTO', to: 'texto', type: 'text' }),
      Object.freeze({ from: 'IMAGEM_URL', to: 'imagemUrl', type: 'url' }),
      Object.freeze({ from: 'ORDEM', to: 'ordem', type: 'number' }),
      Object.freeze({ from: 'ATUALIZADO_EM', to: 'atualizadoEm', type: 'date' })
    ]),
    PORTAL_PUBLIC_PARCEIROS: Object.freeze([
      Object.freeze({ from: 'ID_PARCEIRO', to: 'idParceiro', type: 'text' }),
      Object.freeze({ from: 'NOME', to: 'nome', type: 'text' }),
      Object.freeze({ from: 'TIPO_PARCEIRO', to: 'tipoParceiro', type: 'token' }),
      Object.freeze({ from: 'DESCRICAO', to: 'descricao', type: 'text' }),
      Object.freeze({ from: 'LOGO_URL', to: 'logoUrl', type: 'url' }),
      Object.freeze({ from: 'SITE_URL', to: 'siteUrl', type: 'url' }),
      Object.freeze({ from: 'INSTAGRAM_URL', to: 'instagramUrl', type: 'url' }),
      Object.freeze({ from: 'ORDEM', to: 'ordem', type: 'number' }),
      Object.freeze({ from: 'ATUALIZADO_EM', to: 'atualizadoEm', type: 'date' })
    ]),
    PORTAL_PUBLIC_DOCUMENTOS: Object.freeze([
      Object.freeze({ from: 'ID_DOCUMENTO', to: 'idDocumento', type: 'text' }),
      Object.freeze({ from: 'TITULO', to: 'titulo', type: 'text' }),
      Object.freeze({ from: 'TIPO_DOCUMENTO', to: 'tipoDocumento', type: 'token' }),
      Object.freeze({ from: 'VERSAO', to: 'versao', type: 'text' }),
      Object.freeze({ from: 'DATA_PUBLICACAO', to: 'dataPublicacao', type: 'date' }),
      Object.freeze({ from: 'DESCRICAO', to: 'descricao', type: 'text' }),
      Object.freeze({ from: 'URL_DOCUMENTO', to: 'urlDocumento', type: 'url' }),
      Object.freeze({ from: 'ORDEM', to: 'ordem', type: 'number' }),
      Object.freeze({ from: 'ATUALIZADO_EM', to: 'atualizadoEm', type: 'date' })
    ]),
    PORTAL_PUBLIC_CONFIG: Object.freeze([
      Object.freeze({ from: 'KEY', to: 'key', type: 'token' }),
      Object.freeze({ from: 'VALOR', to: 'valor', type: 'text' }),
      Object.freeze({ from: 'TIPO', to: 'tipo', type: 'token' }),
      Object.freeze({ from: 'GRUPO', to: 'grupo', type: 'token' }),
      Object.freeze({ from: 'DESCRICAO', to: 'descricao', type: 'text' })
    ]),
    PORTAL_PUBLIC_MIDIAS: Object.freeze([
      Object.freeze({ from: 'ID_MIDIA', to: 'idMidia', type: 'text' }),
      Object.freeze({ from: 'NOME', to: 'nome', type: 'text' }),
      Object.freeze({ from: 'TIPO', to: 'tipo', type: 'token' }),
      Object.freeze({ from: 'URL', to: 'url', type: 'url' }),
      Object.freeze({ from: 'DESCRICAO', to: 'descricao', type: 'text' }),
      Object.freeze({ from: 'CATEGORIA', to: 'categoria', type: 'token' }),
      Object.freeze({ from: 'ATUALIZADO_EM', to: 'atualizadoEm', type: 'date' })
    ]),
    PORTAL_PUBLIC_DIRETORIA_COMPLEMENTOS: Object.freeze([
      Object.freeze({ from: 'ID_PESSOA', to: 'ID_PESSOA', type: 'text' }),
      Object.freeze({ from: 'ID_DIRETORIA', to: 'ID_DIRETORIA', type: 'text' }),
      Object.freeze({ from: 'FOTO_URL', to: 'FOTO_URL', type: 'url' }),
      Object.freeze({ from: 'DESCRICAO_PUBLICA', to: 'DESCRICAO_PUBLICA', type: 'text' }),
      Object.freeze({ from: 'LINK_LATTES', to: 'LINK_LATTES', type: 'url' }),
      Object.freeze({ from: 'LINK_INSTAGRAM_PUBLICO', to: 'LINK_INSTAGRAM_PUBLICO', type: 'url' }),
      Object.freeze({ from: 'ORDEM_PUBLICA', to: 'ORDEM_PUBLICA', type: 'number' }),
      Object.freeze({ from: 'STATUS_PUBLICACAO', to: 'STATUS_PUBLICACAO', type: 'token' }),
      Object.freeze({ from: 'PUBLICAR', to: 'PUBLICAR', type: 'token' }),
      Object.freeze({ from: 'ATIVO', to: 'ATIVO', type: 'token' }),
      Object.freeze({ from: 'ATUALIZADO_EM', to: 'ATUALIZADO_EM', type: 'date' })
    ])
  })
});

function corePortalPublicContentCloneDefinitions_() {
  return Object.freeze(CORE_PORTAL_PUBLIC_CONTENT_CFG.definitions.map(function(def) {
    return Object.freeze({
      sheetName: def.sheetName,
      key: def.key,
      headers: Object.freeze(def.headers.slice())
    });
  }));
}

function corePortalPublicContentGetDefinitions_() {
  return Object.freeze({
    spreadsheetName: CORE_PORTAL_PUBLIC_CONTENT_CFG.spreadsheetName,
    registryKeys: CORE_PORTAL_PUBLIC_CONTENT_CFG.registryKeys,
    sheets: corePortalPublicContentCloneDefinitions_(),
    suggestedRegistryRows: Object.freeze(corePortalPublicContentBuildSuggestedRegistryRows_())
  });
}

function corePortalPublicContentBuildSuggestedRegistryRows_(spreadsheetId) {
  var id = String(spreadsheetId || '').trim();
  return CORE_PORTAL_PUBLIC_CONTENT_CFG.definitions.map(function(def) {
    return Object.freeze({
      KEY: def.key,
      SPREADSHEET_ID: id || '<ID_DA_PLANILHA_PORTAL_CONTEUDO_PUBLICO>',
      SHEET_NAME: def.sheetName,
      ATIVO: 'SIM',
      AMBIENTE: 'PROD',
      TYPE: 'SHEET',
      DISPLAY_NAME: CORE_PORTAL_PUBLIC_CONTENT_CFG.spreadsheetName + ' / ' + def.sheetName,
      NOTAS: 'CMS editorial publico do Portal GEAPA'
    });
  });
}

function corePortalPublicContentFindRegisteredSpreadsheetId_() {
  var found = '';
  var missingKeys = [];

  CORE_PORTAL_PUBLIC_CONTENT_CFG.definitions.forEach(function(def) {
    try {
      var meta = core_getRegistryMetaByKey_(def.key);
      if (!found && meta && meta.id) found = String(meta.id || '').trim();
    } catch (err) {
      missingKeys.push(def.key);
    }
  });

  return Object.freeze({
    spreadsheetId: found,
    missingKeys: Object.freeze(missingKeys)
  });
}

function corePortalPublicContentResolveSpreadsheet_(opts) {
  opts = opts || {};
  if (opts.spreadsheet) return opts.spreadsheet;

  var explicitId = String(opts.spreadsheetId || '').trim();
  if (explicitId) return core_openSpreadsheetById_(explicitId);

  var resolved = corePortalPublicContentFindRegisteredSpreadsheetId_();
  if (!resolved.spreadsheetId) {
    throw new Error(
      'Nenhuma KEY de PORTAL_CONTEUDO_PUBLICO encontrada no Registry. ' +
      'Cadastre ao menos PORTAL_PUBLIC_CONFIG ou informe opts.spreadsheetId.'
    );
  }

  return core_openSpreadsheetById_(resolved.spreadsheetId);
}

function corePortalPublicContentCreateSpreadsheet_(options) {
  options = options || {};
  if (options.confirm !== true && options.confirmCreate !== true) {
    throw new Error(
      'Criacao de planilha exige confirmacao explicita: informe { confirm: true }.'
    );
  }

  var name = String(options.name || CORE_PORTAL_PUBLIC_CONTENT_CFG.spreadsheetName).trim();
  var ss = SpreadsheetApp.create(name);

  return Object.freeze({
    ok: true,
    created: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: name,
    suggestedRegistryRows: Object.freeze(corePortalPublicContentBuildSuggestedRegistryRows_(ss.getId()))
  });
}

function corePortalPublicContentGetOrCreateSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    return Object.freeze({
      sheet: sheet,
      created: false
    });
  }

  return Object.freeze({
    sheet: ss.insertSheet(sheetName),
    created: true
  });
}

function corePortalPublicContentEnsureSheetsNoLock_(options) {
  options = options || {};
  var ss = corePortalPublicContentResolveSpreadsheet_(options);
  var actions = [];

  CORE_PORTAL_PUBLIC_CONTENT_CFG.definitions.forEach(function(def) {
    var result = corePortalPublicContentGetOrCreateSheet_(ss, def.sheetName);
    actions.push(Object.freeze({
      sheetName: def.sheetName,
      key: def.key,
      action: result.created ? 'CRIADA' : 'MANTIDA'
    }));
  });

  return Object.freeze({
    ok: true,
    spreadsheetName: typeof ss.getName === 'function' ? ss.getName() : CORE_PORTAL_PUBLIC_CONTENT_CFG.spreadsheetName,
    actions: Object.freeze(actions)
  });
}

function corePortalPublicContentReadHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1 || sheet.getLastRow() < 1) return [];

  return sheet.getRange(CORE_PORTAL_PUBLIC_CONTENT_CFG.headerRow, 1, 1, lastCol)
    .getValues()[0]
    .map(function(value) {
      return String(value || '').trim();
    });
}

function corePortalPublicContentFindMissingHeaders_(existingHeaders, requiredHeaders) {
  var headerMap = core_buildHeaderIndexMap_(existingHeaders, {
    normalize: true,
    oneBased: false,
    keepFirst: true
  });

  return requiredHeaders.filter(function(header) {
    return core_findHeaderIndex_(headerMap, header, {
      normalize: true,
      notFoundValue: -1
    }) < 0;
  });
}

function corePortalPublicContentApplyUx_(sheet, headers) {
  core_freezeHeaderRow_(sheet, CORE_PORTAL_PUBLIC_CONTENT_CFG.headerRow);
  core_ensureFilter_(sheet, CORE_PORTAL_PUBLIC_CONTENT_CFG.headerRow);
  core_applyHeaderColors_(sheet, [
    { color: '#d9ead3', headers: headers.filter(function(header) { return /^ID_|^KEY$/.test(header); }) },
    { color: '#fff2cc', headers: ['STATUS_PUBLICACAO', 'PUBLICAR', 'ATIVO'] },
    { color: '#d0e0e3', headers: ['ORDEM', 'ORDEM_PUBLICA', 'ATUALIZADO_EM', 'DATA_HORA'] }
  ], CORE_PORTAL_PUBLIC_CONTENT_CFG.headerRow);
  core_applyDropdownValidationByHeader_(sheet, {
    PUBLICAR: { values: ['SIM', 'NAO'], allowInvalid: true },
    ATIVO: { values: ['SIM', 'NAO'], allowInvalid: true },
    STATUS_PUBLICACAO: { values: ['RASCUNHO', 'PRONTO', 'PUBLICADO', 'ARQUIVADO'], allowInvalid: true }
  }, CORE_PORTAL_PUBLIC_CONTENT_CFG.headerRow);
}

function corePortalPublicContentEnsureHeadersForSheet_(sheet, def, options) {
  options = options || {};
  var existingHeaders = corePortalPublicContentReadHeaders_(sheet);
  var missing = corePortalPublicContentFindMissingHeaders_(existingHeaders, def.headers);
  var action = 'MANTIDA';

  if (!existingHeaders.length) {
    sheet.getRange(CORE_PORTAL_PUBLIC_CONTENT_CFG.headerRow, 1, 1, def.headers.length)
      .setValues([def.headers.slice()]);
    action = 'CRIADOS_CABECALHOS';
  } else if (missing.length) {
    sheet.getRange(
      CORE_PORTAL_PUBLIC_CONTENT_CFG.headerRow,
      existingHeaders.length + 1,
      1,
      missing.length
    ).setValues([missing]);
    action = 'ADICIONADAS_COLUNAS';
  }

  if (options.applyUx !== false) {
    corePortalPublicContentApplyUx_(sheet, def.headers.slice());
  }

  return Object.freeze({
    sheetName: def.sheetName,
    key: def.key,
    action: action,
    addedHeaders: Object.freeze(missing),
    preservedExtraHeaders: Object.freeze(
      existingHeaders.filter(function(header) {
        return def.headers.indexOf(header) < 0;
      })
    )
  });
}

function corePortalPublicContentEnsureHeadersNoLock_(options) {
  options = options || {};
  var ss = corePortalPublicContentResolveSpreadsheet_(options);
  var actions = [];

  CORE_PORTAL_PUBLIC_CONTENT_CFG.definitions.forEach(function(def) {
    var sheetInfo = corePortalPublicContentGetOrCreateSheet_(ss, def.sheetName);
    var headerResult = corePortalPublicContentEnsureHeadersForSheet_(sheetInfo.sheet, def, options);
    actions.push(Object.freeze({
      sheetName: def.sheetName,
      key: def.key,
      sheetAction: sheetInfo.created ? 'CRIADA' : 'MANTIDA',
      headerAction: headerResult.action,
      addedHeaders: headerResult.addedHeaders,
      preservedExtraHeaders: headerResult.preservedExtraHeaders
    }));
  });

  return Object.freeze({
    ok: true,
    spreadsheetName: typeof ss.getName === 'function' ? ss.getName() : CORE_PORTAL_PUBLIC_CONTENT_CFG.spreadsheetName,
    actions: Object.freeze(actions)
  });
}

function corePortalPublicContentEnsureStructure_(options) {
  options = options || {};
  return core_withLock_(CORE_PORTAL_PUBLIC_CONTENT_CFG.lockKey, function() {
    var result = corePortalPublicContentEnsureHeadersNoLock_(options);
    var resolved = corePortalPublicContentFindRegisteredSpreadsheetId_();
    return Object.freeze({
      ok: true,
      spreadsheetName: result.spreadsheetName,
      actions: result.actions,
      suggestedRegistryRows: Object.freeze(corePortalPublicContentBuildSuggestedRegistryRows_(
        options.spreadsheetId || resolved.spreadsheetId || ''
      ))
    });
  }, Number(options.lockTimeoutMs || 15000));
}

function corePortalPublicContentEnsureSheets_(options) {
  options = options || {};
  return core_withLock_(CORE_PORTAL_PUBLIC_CONTENT_CFG.lockKey, function() {
    return corePortalPublicContentEnsureSheetsNoLock_(options);
  }, Number(options.lockTimeoutMs || 15000));
}

function corePortalPublicContentEnsureHeaders_(options) {
  options = options || {};
  return core_withLock_(CORE_PORTAL_PUBLIC_CONTENT_CFG.lockKey, function() {
    return corePortalPublicContentEnsureHeadersNoLock_(options);
  }, Number(options.lockTimeoutMs || 15000));
}

function corePortalPublicContentDiagnostics_(options) {
  options = options || {};
  var registry = [];
  var resolved = corePortalPublicContentFindRegisteredSpreadsheetId_();
  var ss = null;
  var spreadsheetError = '';

  CORE_PORTAL_PUBLIC_CONTENT_CFG.definitions.forEach(function(def) {
    try {
      var meta = core_getRegistryMetaByKey_(def.key);
      registry.push(Object.freeze({
        key: def.key,
        ok: true,
        spreadsheetId: meta.id,
        sheetName: meta.sheet,
        ativo: meta.ativo,
        ambiente: meta.ambiente,
        lineNo: meta.lineNo
      }));
    } catch (err) {
      registry.push(Object.freeze({
        key: def.key,
        ok: false,
        error: String(err && err.message || err || '')
      }));
    }
  });

  try {
    ss = corePortalPublicContentResolveSpreadsheet_(options);
  } catch (errSpreadsheet) {
    spreadsheetError = String(errSpreadsheet && errSpreadsheet.message || errSpreadsheet || '');
  }

  var sheets = CORE_PORTAL_PUBLIC_CONTENT_CFG.definitions.map(function(def) {
    if (!ss) {
      return Object.freeze({
        sheetName: def.sheetName,
        key: def.key,
        exists: false,
        missingHeaders: Object.freeze(def.headers.slice()),
        error: spreadsheetError
      });
    }

    var sheet = ss.getSheetByName(def.sheetName);
    if (!sheet) {
      return Object.freeze({
        sheetName: def.sheetName,
        key: def.key,
        exists: false,
        missingHeaders: Object.freeze(def.headers.slice())
      });
    }

    var headers = corePortalPublicContentReadHeaders_(sheet);
    return Object.freeze({
      sheetName: def.sheetName,
      key: def.key,
      exists: true,
      missingHeaders: Object.freeze(corePortalPublicContentFindMissingHeaders_(headers, def.headers)),
      extraHeaders: Object.freeze(headers.filter(function(header) {
        return def.headers.indexOf(header) < 0;
      }))
    });
  });

  return Object.freeze({
    ok: registry.every(function(item) { return item.ok; }) &&
      sheets.every(function(item) { return item.exists && item.missingHeaders.length === 0; }),
    spreadsheetName: CORE_PORTAL_PUBLIC_CONTENT_CFG.spreadsheetName,
    registry: Object.freeze(registry),
    sheets: Object.freeze(sheets),
    suggestedRegistryRows: Object.freeze(corePortalPublicContentBuildSuggestedRegistryRows_(
      resolved.spreadsheetId || ''
    ))
  });
}

function corePortalPublicContentBuildError_(errorCode, message, meta) {
  return Object.freeze({
    ok: false,
    errorCode: String(errorCode || 'ERRO_PORTAL_PUBLIC_CONTENT').trim(),
    message: String(message || 'Nao foi possivel ler o conteudo publico do portal.').trim(),
    meta: Object.freeze(meta || {})
  });
}

function corePortalPublicContentNormalizeToken_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  }).replace(/[^A-Z0-9:_-]+/g, '_').replace(/^_+|_+$/g, '');
}

function corePortalPublicContentResolveKey_(keyOrSlug) {
  var raw = String(keyOrSlug || '').trim();
  if (!raw) return '';

  var slug = raw
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .replace(/[\s_-]+/g, '')
    .toLowerCase();

  if (Object.prototype.hasOwnProperty.call(CORE_PORTAL_PUBLIC_CONTENT_READ_CFG.pageSlugMap, slug)) {
    return CORE_PORTAL_PUBLIC_CONTENT_READ_CFG.pageSlugMap[slug];
  }

  return corePortalPublicContentNormalizeToken_(raw);
}

function corePortalPublicContentFindDefinitionByKey_(keyOrSlug) {
  var key = corePortalPublicContentResolveKey_(keyOrSlug);
  for (var i = 0; i < CORE_PORTAL_PUBLIC_CONTENT_CFG.definitions.length; i++) {
    var def = CORE_PORTAL_PUBLIC_CONTENT_CFG.definitions[i];
    if (def.key === key || corePortalPublicContentNormalizeToken_(def.sheetName) === key) {
      return def;
    }
  }
  return null;
}

function corePortalPublicContentHasRecordHeader_(record, headerName) {
  var wanted = core_normalizeHeader_(headerName);
  return Object.keys(record || {}).some(function(key) {
    return core_normalizeHeader_(key) === wanted;
  });
}

function corePortalPublicContentGetRecordValue_(record, headerName) {
  var wanted = core_normalizeHeader_(headerName);
  var keys = Object.keys(record || {});

  for (var i = 0; i < keys.length; i++) {
    if (core_normalizeHeader_(keys[i]) === wanted) {
      return record[keys[i]];
    }
  }

  return '';
}

function corePortalPublicContentIsYes_(value) {
  var token = corePortalPublicContentNormalizeToken_(value);
  return token === 'SIM' || token === 'S' || token === 'TRUE' || token === '1' || token === 'YES';
}

function corePortalPublicContentIsPublishedStatus_(value) {
  var token = corePortalPublicContentNormalizeToken_(value);
  return token === 'PUBLICADO' || token === 'PUBLICADA';
}

function corePortalPublicContentIsPublicableRecord_(record) {
  if (corePortalPublicContentHasRecordHeader_(record, 'ATIVO') &&
      !corePortalPublicContentIsYes_(corePortalPublicContentGetRecordValue_(record, 'ATIVO'))) {
    return false;
  }

  if (corePortalPublicContentHasRecordHeader_(record, 'PUBLICAR') &&
      !corePortalPublicContentIsYes_(corePortalPublicContentGetRecordValue_(record, 'PUBLICAR'))) {
    return false;
  }

  if (corePortalPublicContentHasRecordHeader_(record, 'STATUS_PUBLICACAO') &&
      !corePortalPublicContentIsPublishedStatus_(corePortalPublicContentGetRecordValue_(record, 'STATUS_PUBLICACAO'))) {
    return false;
  }

  return true;
}

function corePortalPublicContentSanitizeText_(value) {
  return String(value == null ? '' : value)
    .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function corePortalPublicContentSanitizeUrl_(value) {
  var url = String(value == null ? '' : value).trim();
  if (!url) return '';
  if (/^(https?:\/\/|\/)/i.test(url)) return url.replace(/[\u0000-\u001F\u007F]/g, '');
  return '';
}

function corePortalPublicContentSanitizeDate_(value) {
  if (!value) return '';
  var parsed = core_parseDateOrNull_(value);
  if (parsed) return core_formatDate_(parsed, null, 'yyyy-MM-dd');
  return corePortalPublicContentSanitizeText_(value);
}

function corePortalPublicContentSanitizeValue_(value, type) {
  if (type === 'url') return corePortalPublicContentSanitizeUrl_(value);
  if (type === 'number') {
    var n = Number(String(value == null ? '' : value).replace(',', '.'));
    return isNaN(n) ? null : n;
  }
  if (type === 'date') return corePortalPublicContentSanitizeDate_(value);
  if (type === 'token') return corePortalPublicContentNormalizeToken_(value);
  return corePortalPublicContentSanitizeText_(value);
}

function corePortalPublicContentBuildPublicItem_(record, key) {
  var fields = CORE_PORTAL_PUBLIC_CONTENT_READ_CFG.publicFields[key] || [];
  var out = {};

  fields.forEach(function(field) {
    out[field.to] = corePortalPublicContentSanitizeValue_(
      corePortalPublicContentGetRecordValue_(record, field.from),
      field.type
    );
  });

  return Object.freeze(out);
}

function corePortalPublicContentResolveOrder_(record) {
  var raw = corePortalPublicContentHasRecordHeader_(record, 'ORDEM_PUBLICA')
    ? corePortalPublicContentGetRecordValue_(record, 'ORDEM_PUBLICA')
    : corePortalPublicContentGetRecordValue_(record, 'ORDEM');
  var n = Number(String(raw == null ? '' : raw).replace(',', '.'));
  return isNaN(n) ? 999999 : n;
}

function corePortalPublicContentSortPublicRecords_(records) {
  return records.slice().sort(function(a, b) {
    var orderA = corePortalPublicContentResolveOrder_(a.rawRecord);
    var orderB = corePortalPublicContentResolveOrder_(b.rawRecord);
    if (orderA !== orderB) return orderA - orderB;
    return String(a.sortLabel || '').localeCompare(String(b.sortLabel || ''), 'pt-BR');
  });
}

function corePortalPublicContentGetCache_(cacheKey, options) {
  if (options && options.disableCache === true) return null;
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    return cached ? JSON.parse(cached) : null;
  } catch (err) {
    return null;
  }
}

function corePortalPublicContentSetCache_(cacheKey, value, options) {
  if (options && options.disableCache === true) return;
  try {
    CacheService.getScriptCache().put(
      cacheKey,
      JSON.stringify(value),
      Number((options && options.cacheTtlSeconds) || CORE_PORTAL_PUBLIC_CONTENT_READ_CFG.cacheTtlSeconds)
    );
  } catch (err) {}
}

function corePortalPublicContentReadRecordsByKey_(key, options) {
  options = options || {};
  if (options.recordsByKey && Array.isArray(options.recordsByKey[key])) {
    return options.recordsByKey[key];
  }

  return core_readRecordsByKey_(key, {
    skipBlankRows: true
  });
}

function corePortalPublicContentReadRows_(keyOrSlug, options) {
  options = options || {};
  var def = corePortalPublicContentFindDefinitionByKey_(keyOrSlug);

  if (!def) {
    return corePortalPublicContentBuildError_(
      'PORTAL_PUBLIC_KEY_DESCONHECIDA',
      'Key de conteudo publico nao reconhecida.',
      { requestedKey: String(keyOrSlug || '').trim() }
    );
  }

  if (CORE_PORTAL_PUBLIC_CONTENT_READ_CFG.blockedPublicKeys.indexOf(def.key) >= 0 &&
      options.includePublicationLog !== true) {
    return corePortalPublicContentBuildError_(
      'PORTAL_PUBLIC_KEY_NAO_EXPONIVEL',
      'Esta aba nao faz parte da leitura publica sanitizada.',
      { key: def.key }
    );
  }

  var cacheKey = CORE_PORTAL_PUBLIC_CONTENT_READ_CFG.cachePrefix + 'ROWS_' + def.key;
  if (!options.recordsByKey && !options.dryRun) {
    var cached = corePortalPublicContentGetCache_(cacheKey, options);
    if (cached) return cached;
  }

  var records;
  try {
    records = corePortalPublicContentReadRecordsByKey_(def.key, options);
  } catch (err) {
    return corePortalPublicContentBuildError_(
      'PORTAL_PUBLIC_REGISTRY_KEY_AUSENTE',
      'Nao foi possivel resolver a key obrigatoria no Registry.',
      { key: def.key, error: String(err && err.message || err || '') }
    );
  }

  var wrapped = records
    .filter(corePortalPublicContentIsPublicableRecord_)
    .map(function(record) {
      var item = corePortalPublicContentBuildPublicItem_(record, def.key);
      return {
        item: item,
        rawRecord: record,
        sortLabel: item.titulo || item.nome || item.idBloco || item.idMarco || item.ID_PESSOA || ''
      };
    });

  var data = corePortalPublicContentSortPublicRecords_(wrapped).map(function(row) {
    return row.item;
  });

  var latest = corePortalPublicContentResolveLatestUpdated_(data);
  var result = Object.freeze({
    ok: true,
    key: def.key,
    sheetName: def.sheetName,
    data: Object.freeze(data),
    meta: Object.freeze({
      total: data.length,
      atualizadoEm: latest,
      origem: 'GEAPA_CORE',
      fonte: CORE_PORTAL_PUBLIC_CONTENT_CFG.spreadsheetName
    })
  });

  if (!options.recordsByKey && !options.dryRun) {
    corePortalPublicContentSetCache_(cacheKey, result, options);
  }

  return result;
}

function corePortalPublicContentResolveLatestUpdated_(items) {
  var latest = '';
  (items || []).forEach(function(item) {
    var value = item.atualizadoEm || item.ATUALIZADO_EM || item.dataPublicacao || item.data || '';
    if (String(value || '') > latest) latest = String(value || '');
  });
  return latest;
}

function corePortalPublicContentReadConfig_(options) {
  options = options || {};
  var key = 'PORTAL_PUBLIC_CONFIG';
  var records;

  try {
    records = corePortalPublicContentReadRecordsByKey_(key, options);
  } catch (err) {
    return corePortalPublicContentBuildError_(
      'PORTAL_PUBLIC_REGISTRY_KEY_AUSENTE',
      'Nao foi possivel resolver a key obrigatoria no Registry.',
      { key: key, error: String(err && err.message || err || '') }
    );
  }

  var config = {};
  records.filter(corePortalPublicContentIsPublicableRecord_).forEach(function(record) {
    var configKey = corePortalPublicContentNormalizeToken_(
      corePortalPublicContentGetRecordValue_(record, 'KEY')
    );
    if (!configKey) return;

    config[configKey] = corePortalPublicContentSanitizeValue_(
      corePortalPublicContentGetRecordValue_(record, 'VALOR'),
      corePortalPublicContentGetRecordValue_(record, 'TIPO') || 'text'
    );
  });

  return Object.freeze({
    ok: true,
    data: Object.freeze(config),
    meta: Object.freeze({
      total: Object.keys(config).length,
      origem: 'GEAPA_CORE',
      fonte: CORE_PORTAL_PUBLIC_CONTENT_CFG.spreadsheetName
    })
  });
}

function corePortalPublicContentGetPage_(slug, options) {
  var normalizedSlug = String(slug || '').trim();
  var key = corePortalPublicContentResolveKey_(normalizedSlug);
  var rows = corePortalPublicContentReadRows_(key, options || {});
  if (!rows.ok) return rows;

  var dataKey = normalizedSlug || key;
  if (key === 'PORTAL_PUBLIC_HISTORIA') dataKey = 'historia';
  if (key === 'PORTAL_PUBLIC_PARCEIROS') dataKey = 'parceiros';

  var listName = 'blocos';
  if (key === 'PORTAL_PUBLIC_HISTORIA') listName = 'marcos';
  if (key === 'PORTAL_PUBLIC_PARCEIROS') listName = 'itens';

  var page = {
    atualizadoEm: rows.meta.atualizadoEm || ''
  };
  page[listName] = rows.data;

  return Object.freeze({
    ok: true,
    slug: dataKey,
    data: Object.freeze(page),
    meta: rows.meta
  });
}

function corePortalPublicContentGetHome_(options) {
  return corePortalPublicContentGetPage_('home', options || {});
}

function corePortalPublicContentGetSobre_(options) {
  return corePortalPublicContentGetPage_('sobre', options || {});
}

function corePortalPublicContentGetHistoria_(options) {
  return corePortalPublicContentGetPage_('historia', options || {});
}

function corePortalPublicContentGetParceiros_(options) {
  return corePortalPublicContentGetPage_('parceiros', options || {});
}

function corePortalPublicContentGetDocumentos_(options) {
  return corePortalPublicContentReadRows_('documentos', options || {});
}

function corePortalPublicContentGetConfig_(options) {
  return corePortalPublicContentReadConfig_(options || {});
}

function corePortalPublicContentGetMidias_(options) {
  return corePortalPublicContentReadRows_('midias', options || {});
}

function corePortalPublicContentGetDiretoriaComplementos_(options) {
  return corePortalPublicContentReadRows_('diretoriaComplementos', options || {});
}

function corePortalPublicContentBuildPublicSnapshot_(options) {
  options = options || {};
  var cacheKey = CORE_PORTAL_PUBLIC_CONTENT_READ_CFG.cachePrefix + 'SNAPSHOT';
  if (!options.recordsByKey && !options.dryRun) {
    var cached = corePortalPublicContentGetCache_(cacheKey, options);
    if (cached) return cached;
  }

  var home = corePortalPublicContentGetHome_(options);
  var sobre = corePortalPublicContentGetSobre_(options);
  var historia = corePortalPublicContentGetHistoria_(options);
  var parceiros = corePortalPublicContentGetParceiros_(options);
  var documentos = corePortalPublicContentGetDocumentos_(options);
  var config = corePortalPublicContentGetConfig_(options);
  var midias = corePortalPublicContentGetMidias_(options);
  var board = corePortalPublicContentGetDiretoriaComplementos_(options);
  var parts = [home, sobre, historia, parceiros, documentos, config, midias, board];
  var errors = parts.filter(function(part) { return !part.ok; });

  if (errors.length) {
    return corePortalPublicContentBuildError_(
      'PORTAL_PUBLIC_SNAPSHOT_INCOMPLETO',
      'Nao foi possivel montar o snapshot publico completo.',
      { errors: errors.map(function(err) { return err.meta || err; }) }
    );
  }

  var updatedValues = [
    home.data.atualizadoEm,
    sobre.data.atualizadoEm,
    historia.data.atualizadoEm,
    parceiros.data.atualizadoEm,
    documentos.meta.atualizadoEm,
    midias.meta.atualizadoEm,
    board.meta.atualizadoEm
  ].filter(Boolean).sort();

  var result = Object.freeze({
    ok: true,
    data: Object.freeze({
      pages: Object.freeze({
        home: home.data,
        sobre: sobre.data,
        historia: historia.data,
        parceiros: parceiros.data
      }),
      documents: documentos.data,
      media: midias.data,
      config: config.data,
      boardComplements: board.data
    }),
    meta: Object.freeze({
      origem: 'GEAPA_CORE',
      fonte: CORE_PORTAL_PUBLIC_CONTENT_CFG.spreadsheetName,
      atualizadoEm: updatedValues.length ? updatedValues[updatedValues.length - 1] : '',
      dryRun: options.dryRun === true
    })
  });

  if (!options.recordsByKey && !options.dryRun) {
    corePortalPublicContentSetCache_(cacheKey, result, options);
  }

  return result;
}

function corePortalPublicContentBuildPublicBoard_(options) {
  return corePortalPublicContentBuildError_(
    'FUNCAO_FUTURA',
    'Montagem de diretoria publica completa fica para etapa futura; use complementos publicos sem cruzar Vigencias/Pessoas.',
    { requested: 'corePortalPublicContentBuildPublicBoard' }
  );
}
