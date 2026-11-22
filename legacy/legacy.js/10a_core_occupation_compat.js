/**
 * Compatibilidade semantica entre "Cargo/Funcao" e "Ocupacao".
 *
 * Nesta etapa, o ecossistema passa a preferir "Ocupacao" em textos novos,
 * mas continua aceitando os cabecalhos legados das planilhas oficiais.
 */

const CORE_OCCUPATION_COMPAT = Object.freeze({
  labels: Object.freeze({
    occupation: 'Ocupa\u00E7\u00E3o',
    currentOccupation: 'Ocupa\u00E7\u00E3o atual'
  }),
  aliases: Object.freeze({
    occupation: Object.freeze([
      'Ocupa\u00E7\u00E3o',
      'Ocupacao',
      'Ocupa\u00E7\u00E3o institucional',
      'Ocupacao institucional',
      'Cargo/Fun\u00E7\u00E3o',
      'Cargo/Funcao',
      'Cargo',
      'FUNCAO',
      'Funcao',
      'Fun\u00E7\u00E3o'
    ]),
    currentOccupation: Object.freeze([
      'Ocupa\u00E7\u00E3o atual',
      'Ocupacao atual',
      'OCUPA\u00C7AO_ATUAL',
      'OCUPACAO_ATUAL',
      'Ocupa\u00E7\u00E3o',
      'Ocupacao',
      'Cargo/Fun\u00E7\u00E3o atual',
      'Cargo/Funcao atual',
      'Cargo/fun\u00E7\u00E3o atual',
      'Cargo/funcao atual',
      'CARGO_FUNCAO_ATUAL',
      'Cargo/Fun\u00E7\u00E3o',
      'Cargo/Funcao'
    ])
  })
});

function core_getOccupationHeaderAliases_(fieldName) {
  var aliases = CORE_OCCUPATION_COMPAT.aliases[fieldName];
  if (!aliases) {
    throw new Error('Campo de ocupacao nao suportado: ' + fieldName);
  }
  return aliases;
}

function core_getOccupationFieldLabel_(fieldName) {
  return CORE_OCCUPATION_COMPAT.labels[fieldName] || CORE_OCCUPATION_COMPAT.labels.occupation;
}

function core_findHeaderIndexByAliasesNormalized_(normalizedHeaders, aliases, normalizer) {
  var normalize = typeof normalizer === 'function'
    ? normalizer
    : function(value) { return core_normalizeHeader_(value); };
  var list = Array.isArray(aliases) ? aliases : [aliases];

  for (var i = 0; i < list.length; i++) {
    var idx = normalizedHeaders.indexOf(normalize(list[i]));
    if (idx !== -1) return idx;
  }

  return -1;
}

function core_findOccupationHeaderIndexInHeaders_(headers, fieldName, normalizer) {
  var normalize = typeof normalizer === 'function'
    ? normalizer
    : function(value) { return core_normalizeHeader_(value); };
  var normalizedHeaders = (headers || []).map(function(header) {
    return normalize(header);
  });

  return core_findHeaderIndexByAliasesNormalized_(
    normalizedHeaders,
    core_getOccupationHeaderAliases_(fieldName),
    normalize
  );
}

function core_getOccupationValueFromRowByHeaders_(row, headers, fieldName, normalizer) {
  var normalize = typeof normalizer === 'function'
    ? normalizer
    : function(value) { return core_normalizeHeader_(value); };
  var normalizedHeaders = (headers || []).map(function(header) {
    return normalize(header);
  });
  var aliases = core_getOccupationHeaderAliases_(fieldName);
  var firstMatchIndex = -1;
  var firstMatchValue = '';

  for (var i = 0; i < aliases.length; i++) {
    var idx = normalizedHeaders.indexOf(normalize(aliases[i]));
    if (idx === -1) continue;

    var currentValue = typeof row[idx] === 'undefined' ? '' : row[idx];
    if (firstMatchIndex < 0) {
      firstMatchIndex = idx;
      firstMatchValue = currentValue;
    }

    if (currentValue !== '' && currentValue !== null && typeof currentValue !== 'undefined') {
      return {
        found: true,
        index: idx,
        value: currentValue
      };
    }
  }

  if (firstMatchIndex >= 0) {
    return {
      found: true,
      index: firstMatchIndex,
      value: firstMatchValue
    };
  }

  return {
    found: false,
    index: -1,
    value: ''
  };
}

function core_getPreferredOccupationHeaderNameFromHeaderMap_(headerMap, fieldName) {
  var aliases = core_getOccupationHeaderAliases_(fieldName);

  for (var i = 0; i < aliases.length; i++) {
    if (core_getCol_(headerMap, aliases[i])) {
      return aliases[i];
    }
  }

  return '';
}

function core_getPreferredOccupationColumnIndexFromHeaderMap_(headerMap, fieldName) {
  var headerName = core_getPreferredOccupationHeaderNameFromHeaderMap_(headerMap, fieldName);
  if (!headerName) return -1;

  var col = core_getCol_(headerMap, headerName);
  return col ? col - 1 : -1;
}

function core_writeOccupationValueByHeaderMap_(sheet, rowNumber, headerMap, fieldName, value) {
  var headerName = core_getPreferredOccupationHeaderNameFromHeaderMap_(headerMap, fieldName);
  if (!headerName) return false;

  return core_writeCellByHeader_(sheet, rowNumber, headerMap, headerName, value, {
    normalize: true
  });
}
