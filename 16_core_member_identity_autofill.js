/***************************************
 * 16_core_member_identity_autofill.js
 ***************************************/

const CORE_MEMBER_IDENTITY_BASE_KEY = 'MEMBERS_ATUAIS';

function core_normalizeIdentityKey_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function core_memberIdentityNormalize_(value) {
  return core_normalizeIdentityKey_(value);
}

function core_memberIdentityGetBaseSheet_() {
  return core_getSheetByKey_(CORE_MEMBER_IDENTITY_BASE_KEY);
}

function core_memberIdentityResolveIndexes_(headers) {
  var headerMap = core_buildHeaderIndexMap_(headers, {
    normalize: true,
    oneBased: false
  });

  function pick(candidates) {
    for (var i = 0; i < candidates.length; i++) {
      var idx = core_findHeaderIndex_(headerMap, candidates[i], {
        normalize: true,
        notFoundValue: -1
      });
      if (idx >= 0) return idx;
    }
    return -1;
  }

  return {
    headerMap: headerMap,
    idxName: pick(['MEMBRO', 'Membro', 'NOME_MEMBRO', 'Nome']),
    idxRga: pick(['RGA']),
    idxEmail: pick(['EMAIL', 'E-mail', 'Email']),
    idxStatus: pick(['Status', 'STATUS_CADASTRAL'])
  };
}

function core_findSheetRowByIdentity_(sheet, identity, opts) {
  core_assertRequired_(sheet, 'sheet');
  core_assertRequired_(identity, 'identity');

  opts = opts || {};
  var headerRow = Number(opts.headerRow || 1);
  var startRow = Number(opts.startRow || (headerRow + 1));

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < startRow || lastCol < 1) return { found: false };

  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  var values = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();
  var idx = core_memberIdentityResolveIndexes_(headers);

  var input = typeof identity === 'object' && identity !== null
    ? identity
    : { rga: identity, email: identity, name: identity };

  var targetRga = core_normalizeIdentityKey_(input.rga || '');
  var targetEmail = core_normalizeIdentityKey_(input.email || '');
  var targetName = core_normalizeIdentityKey_(input.name || identity || '');

  function buildResult(row, index, matchBy) {
    var record = core_rowToObject_(headers, row);
    record.__rowNumber = startRow + index;

    return {
      found: true,
      matchBy: matchBy,
      rowNumber: startRow + index,
      rowValues: row,
      headers: headers,
      headerMap: idx.headerMap,
      record: record,
      name: idx.idxName >= 0 ? String(row[idx.idxName] || '').trim() : '',
      rga: idx.idxRga >= 0 ? String(row[idx.idxRga] || '').trim() : '',
      email: idx.idxEmail >= 0 ? String(row[idx.idxEmail] || '').trim() : '',
      status: idx.idxStatus >= 0 ? String(row[idx.idxStatus] || '').trim() : ''
    };
  }

  if (idx.idxRga >= 0 && targetRga) {
    for (var i = 0; i < values.length; i++) {
      if (core_normalizeIdentityKey_(values[i][idx.idxRga]) === targetRga) {
        return buildResult(values[i], i, 'RGA');
      }
    }
  }

  if (idx.idxEmail >= 0 && targetEmail) {
    for (var j = 0; j < values.length; j++) {
      if (core_normalizeIdentityKey_(values[j][idx.idxEmail]) === targetEmail) {
        return buildResult(values[j], j, 'EMAIL');
      }
    }
  }

  if (idx.idxName >= 0 && targetName) {
    for (var k = 0; k < values.length; k++) {
      if (core_normalizeIdentityKey_(values[k][idx.idxName]) === targetName) {
        return buildResult(values[k], k, 'NOME');
      }
    }
  }

  return {
    found: false,
    headers: headers,
    headerMap: idx.headerMap
  };
}

function core_findMemberCurrentRowByAny_(identity) {
  var sh = core_memberIdentityGetBaseSheet_();
  if (!sh) throw new Error('MEMBERS_ATUAIS nao encontrada.');
  return core_findSheetRowByIdentity_(sh, identity);
}

function core_memberIdentityFindByAny_(identity) {
  core_assertRequired_(identity, 'identity');

  var found = core_findMemberCurrentRowByAny_(identity);
  if (!found || !found.found) return null;

  return {
    found: true,
    matchBy: found.matchBy,
    rowNumber: found.rowNumber,
    name: found.name,
    rga: found.rga,
    email: found.email,
    status: found.status
  };
}

function core_autofillIdentityRowInSheet_(sheet, rowNumber, opts) {
  if (!sheet) throw new Error('sheet obrigatoria');
  if (!rowNumber || rowNumber < 2) throw new Error('rowNumber invalido');

  opts = opts || {};

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { ok: false, reason: 'sheet vazia' };

  var headerMap = core_headerMap_(sheet, 1);

  var nameHeaders = Array.isArray(opts.nameHeaders) && opts.nameHeaders.length
    ? opts.nameHeaders
    : ['Nome'];
  var rgaHeaders = Array.isArray(opts.rgaHeaders) && opts.rgaHeaders.length
    ? opts.rgaHeaders
    : ['RGA'];
  var emailHeaders = Array.isArray(opts.emailHeaders) && opts.emailHeaders.length
    ? opts.emailHeaders
    : ['E-mail', 'EMAIL'];

  var resolvedName = core_findFirstExistingHeader_(headerMap, nameHeaders);
  var resolvedRga = core_findFirstExistingHeader_(headerMap, rgaHeaders);
  var resolvedEmail = core_findFirstExistingHeader_(headerMap, emailHeaders);

  var colName = resolvedName && resolvedName.found ? resolvedName.index : -1;
  var colRga = resolvedRga && resolvedRga.found ? resolvedRga.index : -1;
  var colEmail = resolvedEmail && resolvedEmail.found ? resolvedEmail.index : -1;

  if (colName < 1 || colRga < 1 || colEmail < 1) {
    throw new Error(
      'A sheet destino precisa ter os cabecalhos ' +
      nameHeaders.join(' / ') + ', ' +
      rgaHeaders.join(' / ') + ' e ' +
      emailHeaders.join(' / ') + '.'
    );
  }

  var row = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];

  var currentName = String(row[colName - 1] || '').trim();
  var currentRga = String(row[colRga - 1] || '').trim();
  var currentEmail = String(row[colEmail - 1] || '').trim();

  var identity = currentRga || currentEmail || currentName;
  if (!identity) return { ok: false, reason: 'linha sem identificador' };

  var member = core_memberIdentityFindByAny_(identity);
  if (!member || !member.found) {
    return { ok: false, reason: 'membro nao encontrado', identity: identity };
  }

  if (currentName && member.name && core_memberIdentityNormalize_(currentName) !== core_memberIdentityNormalize_(member.name)) {
    return { ok: false, reason: 'conflito nome', identity: identity, found: member };
  }
  if (currentRga && member.rga && core_memberIdentityNormalize_(currentRga) !== core_memberIdentityNormalize_(member.rga)) {
    return { ok: false, reason: 'conflito rga', identity: identity, found: member };
  }
  if (currentEmail && member.email && core_memberIdentityNormalize_(currentEmail) !== core_memberIdentityNormalize_(member.email)) {
    return { ok: false, reason: 'conflito email', identity: identity, found: member };
  }

  var updated = [];

  if (!currentName && member.name) {
    sheet.getRange(rowNumber, colName).setValue(member.name);
    updated.push('Nome');
  }

  if (!currentRga && member.rga) {
    sheet.getRange(rowNumber, colRga).setValue(member.rga);
    updated.push('RGA');
  }

  if (!currentEmail && member.email) {
    sheet.getRange(rowNumber, colEmail).setValue(member.email);
    updated.push(emailHeaders[0]);
  }

  return {
    ok: true,
    rowNumber: rowNumber,
    matchBy: member.matchBy,
    updated: updated,
    found: member
  };
}
