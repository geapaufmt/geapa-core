/***************************************
 * 16b_core_entity_identity_ids.js
 *
 * Identidade institucional complementar do ecossistema GEAPA.
 *
 * Regras:
 * - membros continuam identificados por RGA;
 * - professores usam ID_PROFESSOR;
 * - participantes externos usam ID_PARTICIPANTE_EXTERNO;
 * - IDs novos sao gerados apenas quando a celula estiver vazia;
 * - IDs existentes nunca sao recalculados;
 * - externos validam duplicidade por e-mail antes de criar novo ID.
 ***************************************/

const CORE_ENTITY_ID_CFG = Object.freeze({
  professors: Object.freeze({
    entityKey: 'professors',
    registryKey: 'PROFS_BASE',
    lockKey: 'CORE_PROFESSOR_IDS',
    idHeaders: Object.freeze(['ID_PROFESSOR']),
    emailHeaders: Object.freeze(['EMAIL', 'Email', 'E-mail']),
    nameHeaders: Object.freeze(['Professor', 'PROFESSOR', 'Nome', 'NOME']),
    prefix: 'PROF-',
    padLength: 4,
    duplicateByEmail: false,
    requireEmailForAutoId: false
  }),

  externals: Object.freeze({
    entityKey: 'externals',
    registryKey: 'PARTICIPANTES_EXTERNOS_BASE',
    lockKey: 'CORE_EXTERNAL_IDS',
    idHeaders: Object.freeze(['ID_PARTICIPANTE_EXTERNO']),
    emailHeaders: Object.freeze(['EMAIL', 'Email', 'E-mail']),
    nameHeaders: Object.freeze(['Nome', 'NOME', 'Participante', 'PARTICIPANTE']),
    prefix: 'EXT-',
    padLength: 4,
    duplicateByEmail: true,
    requireEmailForAutoId: true
  })
});

function core_identityGetConfig_(entityType) {
  var key = String(entityType || '').trim().toLowerCase();
  if (!CORE_ENTITY_ID_CFG[key]) {
    throw new Error('Tipo de identidade nao suportado: ' + entityType);
  }
  return CORE_ENTITY_ID_CFG[key];
}

function core_identityPickIndex_(headerMap, aliases) {
  var names = Array.isArray(aliases) ? aliases : [aliases];

  for (var i = 0; i < names.length; i++) {
    var idx = core_findHeaderIndex_(headerMap, names[i], {
      normalize: true,
      notFoundValue: -1
    });
    if (idx >= 0) return idx;
  }

  return -1;
}

function core_identityResolveIndexes_(headers, cfg) {
  var headerMap = core_buildHeaderIndexMap_(headers, {
    normalize: true,
    oneBased: false,
    keepFirst: true
  });

  return Object.freeze({
    headerMap: headerMap,
    idxId: core_identityPickIndex_(headerMap, cfg.idHeaders),
    idxEmail: core_identityPickIndex_(headerMap, cfg.emailHeaders),
    idxName: core_identityPickIndex_(headerMap, cfg.nameHeaders)
  });
}

function core_identityReadContext_(cfg) {
  var sheet = core_getSheetByKey_(cfg.registryKey);
  var data = core_readSheetData_(sheet, {
    headerRow: 1,
    startRow: 2,
    normalizeHeaderMap: false
  });
  var idx = core_identityResolveIndexes_(data.headers, cfg);

  if (idx.idxId < 0) {
    throw new Error(
      'Cabecalho obrigatorio nao encontrado em ' + cfg.registryKey + ': ' + cfg.idHeaders.join(' / ')
    );
  }

  if (cfg.requireEmailForAutoId && idx.idxEmail < 0) {
    throw new Error(
      'Cabecalho obrigatorio nao encontrado em ' + cfg.registryKey + ': ' + cfg.emailHeaders.join(' / ')
    );
  }

  return {
    cfg: cfg,
    sheet: sheet,
    headers: data.headers,
    rows: data.rows,
    idx: idx,
    startRow: 2,
    lastCol: data.headers.length
  };
}

function core_identityExtractSequenceNumber_(idValue, prefix) {
  var text = String(idValue || '').trim().toUpperCase();
  if (!text) return 0;

  var escapedPrefix = String(prefix || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&').toUpperCase();
  var match = text.match(new RegExp('^' + escapedPrefix + '(\\d+)$'));
  if (!match) return 0;

  return parseInt(match[1], 10) || 0;
}

function core_identityFormatId_(prefix, sequence, padLength) {
  var padded = String(Number(sequence || 0)).padStart(Number(padLength || 4), '0');
  return String(prefix || '') + padded;
}

function core_identityNormalizeOptionalEmail_(value) {
  var email = core_normalizeEmail_(value);
  return core_isValidEmail_(email) ? email : '';
}

function core_identityBuildSnapshot_(ctx, rowValues, zeroBasedIndex) {
  var rowNumber = ctx.startRow + zeroBasedIndex;
  var idValue = ctx.idx.idxId >= 0 ? String(rowValues[ctx.idx.idxId] || '').trim() : '';
  var emailValue = ctx.idx.idxEmail >= 0 ? core_identityNormalizeOptionalEmail_(rowValues[ctx.idx.idxEmail]) : '';
  var nameValue = ctx.idx.idxName >= 0 ? String(rowValues[ctx.idx.idxName] || '').trim() : '';

  return {
    rowNumber: rowNumber,
    id: idValue,
    email: emailValue,
    name: nameValue,
    values: rowValues
  };
}

function core_identityToPublicSnapshot_(snapshot) {
  return Object.freeze({
    rowNumber: snapshot.rowNumber,
    id: snapshot.id,
    email: snapshot.email,
    name: snapshot.name
  });
}

function core_identityCollectState_(ctx) {
  var maxSequence = 0;
  var malformedIds = [];
  var snapshots = [];
  var byEmail = Object.create(null);

  for (var i = 0; i < ctx.rows.length; i++) {
    var snapshot = core_identityBuildSnapshot_(ctx, ctx.rows[i], i);
    snapshots.push(snapshot);

    if (snapshot.id) {
      var sequence = core_identityExtractSequenceNumber_(snapshot.id, ctx.cfg.prefix);
      if (sequence > 0) {
        if (sequence > maxSequence) maxSequence = sequence;
      } else {
        malformedIds.push({ rowNumber: snapshot.rowNumber, id: snapshot.id });
      }
    }

    if (ctx.cfg.duplicateByEmail && snapshot.email) {
      if (!byEmail[snapshot.email]) {
        byEmail[snapshot.email] = { email: snapshot.email, rows: [] };
      }
      byEmail[snapshot.email].rows.push(snapshot);
    }
  }

  return {
    maxSequence: maxSequence,
    malformedIds: malformedIds,
    snapshots: snapshots,
    byEmail: byEmail
  };
}

function core_identityTakeNextId_(state, cfg) {
  state.maxSequence += 1;
  return core_identityFormatId_(cfg.prefix, state.maxSequence, cfg.padLength);
}

function core_identityWriteId_(ctx, rowNumber, idValue) {
  ctx.sheet.getRange(rowNumber, ctx.idx.idxId + 1).setValue(idValue);
  return idValue;
}

function core_identityBuildDuplicateSummary_(emailGroup) {
  var group = emailGroup || { rows: [] };
  var ids = group.rows
    .map(function(snapshot) { return snapshot.id; })
    .filter(function(idValue) { return !!idValue; });

  return Object.freeze({
    email: group.email || '',
    rowNumbers: group.rows.map(function(snapshot) { return snapshot.rowNumber; }),
    ids: ids,
    matches: group.rows.map(core_identityToPublicSnapshot_)
  });
}

function core_identityGetSnapshotByRowNumber_(state, rowNumber) {
  for (var i = 0; i < state.snapshots.length; i++) {
    if (state.snapshots[i].rowNumber === rowNumber) return state.snapshots[i];
  }
  return null;
}

function core_identityFindExternalByEmail_(email) {
  core_assertRequired_(email, 'email');

  var cfg = core_identityGetConfig_('externals');
  var ctx = core_identityReadContext_(cfg);
  var state = core_identityCollectState_(ctx);
  var normalizedEmail = core_identityNormalizeOptionalEmail_(email);
  if (!normalizedEmail) {
    return Object.freeze({ found: false, duplicate: false, email: '' });
  }

  var emailGroup = state.byEmail[normalizedEmail];
  if (!emailGroup || !emailGroup.rows.length) {
    return Object.freeze({ found: false, duplicate: false, email: normalizedEmail });
  }

  var rowsWithId = emailGroup.rows.filter(function(snapshot) { return !!snapshot.id; });
  var canonical = rowsWithId.length === 1
    ? rowsWithId[0]
    : (emailGroup.rows.length === 1 ? emailGroup.rows[0] : null);

  return Object.freeze({
    found: true,
    duplicate: emailGroup.rows.length > 1,
    email: normalizedEmail,
    canonical: canonical ? core_identityToPublicSnapshot_(canonical) : null,
    matches: emailGroup.rows.map(core_identityToPublicSnapshot_)
  });
}

function core_identityValidateExternalEmailDuplicates_() {
  var cfg = core_identityGetConfig_('externals');
  var ctx = core_identityReadContext_(cfg);
  var state = core_identityCollectState_(ctx);
  var duplicates = [];

  Object.keys(state.byEmail).forEach(function(email) {
    var emailGroup = state.byEmail[email];
    if (emailGroup.rows.length > 1) {
      duplicates.push(core_identityBuildDuplicateSummary_(emailGroup));
    }
  });

  return Object.freeze({
    ok: duplicates.length === 0,
    duplicateCount: duplicates.length,
    duplicates: duplicates
  });
}

function core_identityEnsureIdForRowByConfig_(cfg, rowNumber) {
  core_assertRequired_(rowNumber, 'rowNumber');

  return core_withLock_(cfg.lockKey, function() {
    var ctx = core_identityReadContext_(cfg);
    var state = core_identityCollectState_(ctx);
    var targetRow = Number(rowNumber || 0);
    var snapshot = core_identityGetSnapshotByRowNumber_(state, targetRow);

    if (!snapshot) {
      return Object.freeze({
        ok: false,
        created: false,
        reason: 'row_not_found',
        rowNumber: targetRow
      });
    }

    if (snapshot.id) {
      return Object.freeze({
        ok: true,
        created: false,
        rowNumber: snapshot.rowNumber,
        id: snapshot.id,
        email: snapshot.email,
        name: snapshot.name
      });
    }

    if (cfg.requireEmailForAutoId && !snapshot.email) {
      return Object.freeze({
        ok: false,
        created: false,
        reason: 'email_required',
        rowNumber: snapshot.rowNumber,
        name: snapshot.name
      });
    }

    if (cfg.duplicateByEmail && snapshot.email) {
      var emailGroup = state.byEmail[snapshot.email];
      if (emailGroup && emailGroup.rows.length > 1) {
        var duplicateSummary = core_identityBuildDuplicateSummary_(emailGroup);
        return Object.freeze({
          ok: false,
          created: false,
          reason: 'duplicate_email',
          rowNumber: snapshot.rowNumber,
          email: snapshot.email,
          duplicate: duplicateSummary,
          existingId: duplicateSummary.ids.length === 1 ? duplicateSummary.ids[0] : ''
        });
      }
    }

    var nextId = core_identityTakeNextId_(state, cfg);
    core_identityWriteId_(ctx, snapshot.rowNumber, nextId);

    return Object.freeze({
      ok: true,
      created: true,
      rowNumber: snapshot.rowNumber,
      id: nextId,
      email: snapshot.email,
      name: snapshot.name
    });
  });
}

function core_identityFillMissingIdsForConfig_(cfg) {
  return core_withLock_(cfg.lockKey, function() {
    var ctx = core_identityReadContext_(cfg);
    var state = core_identityCollectState_(ctx);
    var counters = {
      scanned: state.snapshots.length,
      alreadyFilled: 0,
      created: 0,
      skippedNoEmail: 0,
      duplicatesSkipped: 0,
      malformedExistingIds: state.malformedIds.length
    };
    var createdRows = [];
    var duplicateGroups = [];
    var seenDuplicateEmail = Object.create(null);

    state.snapshots.forEach(function(snapshot) {
      if (snapshot.id) {
        counters.alreadyFilled++;
        return;
      }

      if (cfg.requireEmailForAutoId && !snapshot.email) {
        counters.skippedNoEmail++;
        return;
      }

      if (cfg.duplicateByEmail && snapshot.email) {
        var emailGroup = state.byEmail[snapshot.email];
        if (emailGroup && emailGroup.rows.length > 1) {
          counters.duplicatesSkipped++;
          if (!seenDuplicateEmail[snapshot.email]) {
            duplicateGroups.push(core_identityBuildDuplicateSummary_(emailGroup));
            seenDuplicateEmail[snapshot.email] = true;
          }
          return;
        }
      }

      var nextId = core_identityTakeNextId_(state, cfg);
      core_identityWriteId_(ctx, snapshot.rowNumber, nextId);
      snapshot.id = nextId;
      counters.created++;
      createdRows.push(core_identityToPublicSnapshot_(snapshot));
    });

    return Object.freeze({
      ok: true,
      registryKey: cfg.registryKey,
      counters: Object.freeze(counters),
      createdRows: createdRows,
      duplicateGroups: duplicateGroups,
      malformedExistingIds: state.malformedIds
    });
  });
}

function core_fillMissingProfessorIds_() {
  return core_identityFillMissingIdsForConfig_(core_identityGetConfig_('professors'));
}

function core_fillMissingExternalIds_() {
  return core_identityFillMissingIdsForConfig_(core_identityGetConfig_('externals'));
}

function core_ensureProfessorIdForRow_(rowNumber) {
  return core_identityEnsureIdForRowByConfig_(core_identityGetConfig_('professors'), rowNumber);
}

function core_ensureExternalIdForRow_(rowNumber) {
  return core_identityEnsureIdForRowByConfig_(core_identityGetConfig_('externals'), rowNumber);
}
