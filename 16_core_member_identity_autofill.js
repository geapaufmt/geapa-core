/***************************************
 * 16_core_member_identity_autofill.js
 ***************************************/

const CORE_MEMBER_IDENTITY_BASE_KEY = 'MEMBERS_ATUAIS';

function core_memberIdentityNormalize_(value) {
  return String(value == null ? '' : value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function core_memberIdentityGetBaseSheet_() {
  return core_getSheetByKey_(CORE_MEMBER_IDENTITY_BASE_KEY);
}

function core_memberIdentityFindByAny_(identity) {
  core_assertRequired_(identity, 'identity');

  var sh = core_memberIdentityGetBaseSheet_();
  if (!sh) throw new Error('MEMBERS_ATUAIS não encontrada.');

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  var headerMap = core_headerMap_(sh, 1);
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var idxName = core_getCol_(headerMap, 'MEMBRO') || core_getCol_(headerMap, 'Nome');
  var idxRga = core_getCol_(headerMap, 'RGA');
  var idxEmail = core_getCol_(headerMap, 'EMAIL') || core_getCol_(headerMap, 'E-mail');
  var idxStatus = core_getCol_(headerMap, 'Status');

  idxName = idxName ? idxName - 1 : -1;
  idxRga = idxRga ? idxRga - 1 : -1;
  idxEmail = idxEmail ? idxEmail - 1 : -1;
  idxStatus = idxStatus ? idxStatus - 1 : -1;

  var target = core_memberIdentityNormalize_(identity);

  for (var i = 0; i < values.length; i++) {
    if (idxRga >= 0 && core_memberIdentityNormalize_(values[i][idxRga]) === target) {
      return {
        found: true,
        matchBy: 'RGA',
        rowNumber: i + 2,
        name: idxName >= 0 ? String(values[i][idxName] || '').trim() : '',
        rga: idxRga >= 0 ? String(values[i][idxRga] || '').trim() : '',
        email: idxEmail >= 0 ? String(values[i][idxEmail] || '').trim() : '',
        status: idxStatus >= 0 ? String(values[i][idxStatus] || '').trim() : ''
      };
    }
  }

  for (var j = 0; j < values.length; j++) {
    if (idxEmail >= 0 && core_memberIdentityNormalize_(values[j][idxEmail]) === target) {
      return {
        found: true,
        matchBy: 'EMAIL',
        rowNumber: j + 2,
        name: idxName >= 0 ? String(values[j][idxName] || '').trim() : '',
        rga: idxRga >= 0 ? String(values[j][idxRga] || '').trim() : '',
        email: idxEmail >= 0 ? String(values[j][idxEmail] || '').trim() : '',
        status: idxStatus >= 0 ? String(values[j][idxStatus] || '').trim() : ''
      };
    }
  }

  for (var k = 0; k < values.length; k++) {
    if (idxName >= 0 && core_memberIdentityNormalize_(values[k][idxName]) === target) {
      return {
        found: true,
        matchBy: 'NOME',
        rowNumber: k + 2,
        name: idxName >= 0 ? String(values[k][idxName] || '').trim() : '',
        rga: idxRga >= 0 ? String(values[k][idxRga] || '').trim() : '',
        email: idxEmail >= 0 ? String(values[k][idxEmail] || '').trim() : '',
        status: idxStatus >= 0 ? String(values[k][idxStatus] || '').trim() : ''
      };
    }
  }

  return null;
}

function core_autofillIdentityRowInSheet_(sheet, rowNumber) {
  if (!sheet) throw new Error('sheet obrigatória');
  if (!rowNumber || rowNumber < 2) throw new Error('rowNumber inválido');

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { ok: false, reason: 'sheet vazia' };

  var headerMap = core_headerMap_(sheet, 1);

  var colName = core_getCol_(headerMap, 'Nome');
  var colRga = core_getCol_(headerMap, 'RGA');
  var colEmail = core_getCol_(headerMap, 'E-mail') || core_getCol_(headerMap, 'EMAIL');

  if (!colName || !colRga || !colEmail) {
    throw new Error('A sheet destino precisa ter os cabeçalhos Nome, RGA e E-mail.');
  }

  var row = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];

  var currentName = String(row[colName - 1] || '').trim();
  var currentRga = String(row[colRga - 1] || '').trim();
  var currentEmail = String(row[colEmail - 1] || '').trim();

  var identity = currentRga || currentEmail || currentName;
  if (!identity) return { ok: false, reason: 'linha sem identificador' };

  var member = core_memberIdentityFindByAny_(identity);
  if (!member || !member.found) {
    return { ok: false, reason: 'membro não encontrado', identity: identity };
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
    updated.push('E-mail');
  }

  return {
    ok: true,
    rowNumber: rowNumber,
    matchBy: member.matchBy,
    updated: updated,
    found: member
  };
}