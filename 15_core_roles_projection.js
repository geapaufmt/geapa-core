/***************************************
 * 15_core_roles_projection.js
 *
 * Projeção operacional de cargos atuais
 ***************************************/

const CORE_ROLE_PROJECTION_CURRENT_MEMBERS_KEY = 'MEMBERS_ATUAIS';

const CORE_ROLE_PROJECTION_VIGENCIA_KEYS = Object.freeze({
  diretoria: 'VIGENCIA_MEMBROS_DIRETORIAS',
  assessoria: 'VIGENCIA_ASSESSORES',
  conselho: 'VIGENCIA_CONSELHEIROS'
});

const CORE_ROLE_PROJECTION_HEADERS = Object.freeze({
  memberName: Object.freeze(['MEMBRO', 'Membro', 'NOME_MEMBRO', 'Nome']),
  email: Object.freeze(['EMAIL', 'E-mail', 'Email']),
  rga: Object.freeze(['RGA']),
  phone: Object.freeze(['TELEFONE', 'Telefone']),
  currentOccupation: core_getOccupationHeaderAliases_('currentOccupation'),
  institutionalOccupation: core_getOccupationHeaderAliases_('occupation'),
  status: Object.freeze(['Status', 'STATUS_CADASTRAL']),
  integratedAt: Object.freeze(['Data integra\u00E7\u00E3o', 'Data integracao', 'DATA_INTEGRACAO']),
  entrySemester: Object.freeze(['Semestre de Entrada', 'Semestre de entrada', 'SEMESTRE_ENTRADA']),
  currentSemester: Object.freeze(['Semestre atual', 'SEMESTRE_ATUAL']),
  groupSemesters: Object.freeze(['N\u00B0 de semestres no grupo', 'N\u00BA de semestres no grupo', 'NÂ° de semestres no grupo', 'NÂº de semestres no grupo', 'QTD_SEMESTRES_NO_GRUPO']),
  effectiveGroupTime: Object.freeze(['TEMPO_EFETIVO_NO_GRUPO'])
});

function core_roleProjectionNormalizeKey_(value) {
  return String(value == null ? '' : value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function core_roleProjectionResolveRefDate_(refDate) {
  return refDate instanceof Date && !isNaN(refDate) ? refDate : new Date();
}

function core_roleProjectionGuessColumnIndex_(headerMap, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var col = core_getCol_(headerMap, candidates[i]);
    if (col) return col - 1;
  }
  return -1;
}

function core_roleProjectionParseDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return value;
  var d = new Date(value);
  return isNaN(d) ? null : d;
}

function core_getCivilGroupTimeFromIntegrationDate_(integrationDate, refDate) {
  var start = core_parseDateOrNull_(integrationDate);
  var end = core_parseDateOrNull_(refDate) || new Date();

  if (!start || !end) return null;

  var totalMonths =
    (end.getFullYear() - start.getFullYear()) * 12 +
    (end.getMonth() - start.getMonth());

  if (end.getDate() < start.getDate()) {
    totalMonths -= 1;
  }

  if (totalMonths < 0) totalMonths = 0;

  return Object.freeze({
    totalMonths: totalMonths,
    years: Math.floor(totalMonths / 12),
    months: totalMonths % 12
  });
}

function core_formatCivilGroupTime_(civilTime) {
  if (!civilTime) return '';

  var years = Number(civilTime.years || 0);
  var months = Number(civilTime.months || 0);
  var yearsLabel = years === 1 ? 'ano' : 'anos';
  var monthsLabel = months === 1 ? 'mês' : 'meses';

  return years + ' ' + yearsLabel + ' e ' + months + ' ' + monthsLabel;
}

function core_roleProjectionIsAssignmentActive_(startValue, endValue, refDate) {
  var ref = core_startOfDay_(refDate);
  var start = core_roleProjectionParseDate_(startValue);
  var end = core_roleProjectionParseDate_(endValue);

  if (!start && !end) return true;
  if (start && core_startOfDay_(start) > ref) return false;
  if (end && core_startOfDay_(end) < ref) return false;

  return true;
}

function core_roleProjectionReadAssignmentsFromSheet_(sheetKey, sourceType, refDate) {
  var sh = core_getSheetByKey_(sheetKey);
  if (!sh) throw new Error('Sheet não encontrada para key: ' + sheetKey);

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var headerMap = core_headerMap_(sh, 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(value) {
    return String(value || '').trim();
  });
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var idxName = core_roleProjectionGuessColumnIndex_(headerMap, ['Nome', 'MEMBRO']);
  var idxEmail = core_roleProjectionGuessColumnIndex_(headerMap, ['E-mail', 'EMAIL']);
  var idxRga = core_roleProjectionGuessColumnIndex_(headerMap, ['RGA']);
  var idxStart = core_roleProjectionGuessColumnIndex_(headerMap, ['Data_Início', 'Data início', 'Data de início', 'INICIO', 'Início']);
  var idxEndReal = core_roleProjectionGuessColumnIndex_(headerMap, ['Data_Fim', 'Data fim', 'Data de fim', 'FIM', 'Fim']);
  var idxEndPlanned = core_roleProjectionGuessColumnIndex_(headerMap, ['Data_Fim_previsto', 'Data fim previsto', 'Data de fim prevista', 'Fim previsto']);
  var idxAtivo = core_roleProjectionGuessColumnIndex_(headerMap, ['ATIVO', 'Ativo']);

  var out = [];

  values.forEach(function(row) {
    var memberName = idxName >= 0 ? String(row[idxName] || '').trim() : '';
    var email = idxEmail >= 0 ? String(row[idxEmail] || '').trim() : '';
    var rga = idxRga >= 0 ? String(row[idxRga] || '').trim() : '';
    var occupationCell = core_getOccupationValueFromRowByHeaders_(
      row,
      headers,
      'occupation',
      core_roleProjectionNormalizeKey_
    );
    var rawOccupation = occupationCell.found ? String(occupationCell.value || '').trim() : '';

    if (!memberName && !email && !rga) return;
    if (!rawOccupation) return;

    if (idxAtivo >= 0) {
      var ativo = core_rolesParseYesNo_(row[idxAtivo]);
      if (!ativo) return;
    }

    var startValue = idxStart >= 0 ? row[idxStart] : null;
    var endRealValue = idxEndReal >= 0 ? row[idxEndReal] : null;
    var endPlannedValue = idxEndPlanned >= 0 ? row[idxEndPlanned] : null;
    var effectiveEndValue = endRealValue || endPlannedValue;

    if (!core_roleProjectionIsAssignmentActive_(startValue, effectiveEndValue, refDate)) return;

    var roleConfig = core_findInstitutionalRoleByAnyName_(rawOccupation);

    out.push(Object.freeze({
      sourceType: sourceType,
      rawOccupation: rawOccupation,
      rawRole: rawOccupation,
      roleConfig: roleConfig,
      roleKey: roleConfig ? roleConfig.roleKey : null,
      occupation: roleConfig ? roleConfig.publicName : rawOccupation,
      occupationName: roleConfig ? roleConfig.publicName : rawOccupation,
      publicName: roleConfig ? roleConfig.publicName : rawOccupation,
      displayOrder: roleConfig ? roleConfig.displayOrder : 999999,
      receivesEmails: roleConfig ? roleConfig.receivesEmails : false,
      emailGroups: roleConfig ? roleConfig.emailGroups : Object.freeze([]),
      isUniqueRole: roleConfig ? roleConfig.isUniqueRole : false,
      memberName: memberName,
      memberNameNorm: core_roleProjectionNormalizeKey_(memberName),
      email: email,
      emailNorm: core_roleProjectionNormalizeKey_(email),
      rga: rga,
      rgaNorm: core_roleProjectionNormalizeKey_(rga),
      startDate: core_roleProjectionParseDate_(startValue),
      endDateReal: core_roleProjectionParseDate_(endRealValue),
      endDatePlanned: core_roleProjectionParseDate_(endPlannedValue),
      endDateEffective: core_roleProjectionParseDate_(effectiveEndValue)
    }));
  });

  return out;
}

function core_getCurrentInstitutionalAssignments_(refDate) {
  var ref = core_roleProjectionResolveRefDate_(refDate);

  var diretoria = core_roleProjectionReadAssignmentsFromSheet_(
    CORE_ROLE_PROJECTION_VIGENCIA_KEYS.diretoria,
    'DIRETORIA',
    ref
  );

  var assessoria = core_roleProjectionReadAssignmentsFromSheet_(
    CORE_ROLE_PROJECTION_VIGENCIA_KEYS.assessoria,
    'ASSESSORIA',
    ref
  );

  var conselho = core_roleProjectionReadAssignmentsFromSheet_(
    CORE_ROLE_PROJECTION_VIGENCIA_KEYS.conselho,
    'CONSELHO',
    ref
  );

  return Object.freeze([].concat(diretoria, assessoria, conselho));
}

function core_roleProjectionBuildMemberIdentityKey_(assignment) {
  if (assignment.rgaNorm) return 'RGA::' + assignment.rgaNorm;
  if (assignment.emailNorm) return 'EMAIL::' + assignment.emailNorm;
  return 'NOME::' + assignment.memberNameNorm;
}

function core_groupCurrentInstitutionalAssignmentsByMember_(refDate) {
  var rows = core_getCurrentInstitutionalAssignments_(refDate);
  var grouped = {};

  rows.forEach(function(item) {
    var key = core_roleProjectionBuildMemberIdentityKey_(item);
    if (!grouped[key]) {
      grouped[key] = {
        identityKey: key,
        memberName: item.memberName,
        memberNameNorm: item.memberNameNorm,
        email: item.email,
        emailNorm: item.emailNorm,
        rga: item.rga,
        rgaNorm: item.rgaNorm,
        roles: []
      };
    }
    grouped[key].roles.push(item);
  });

  Object.keys(grouped).forEach(function(k) {
    grouped[k].roles.sort(function(a, b) {
      return a.displayOrder - b.displayOrder;
    });
    grouped[k] = Object.freeze({
      identityKey: grouped[k].identityKey,
      memberName: grouped[k].memberName,
      memberNameNorm: grouped[k].memberNameNorm,
      email: grouped[k].email,
      emailNorm: grouped[k].emailNorm,
      rga: grouped[k].rga,
      rgaNorm: grouped[k].rgaNorm,
      roles: Object.freeze(grouped[k].roles.slice())
    });
  });

  return Object.freeze(grouped);
}

function core_buildCurrentOccupationCellValue_(roles) {
  if (!roles || !roles.length) return 'Membro';

  var uniqueNames = [];
  var seen = {};

  roles.forEach(function(role) {
    var key = core_roleProjectionNormalizeKey_(role.publicName);
    if (!seen[key]) {
      seen[key] = true;
      uniqueNames.push(role.publicName);
    }
  });

  return uniqueNames.join(' | ');
}

function core_buildCurrentRoleCellValue_(roles) {
  return core_buildCurrentOccupationCellValue_(roles);
}

function core_getCurrentAssignmentsByEmailGroup_(groupName, refDate) {
  var groupNorm = core_rolesNormalizeText_(groupName);
  var rows = core_getCurrentInstitutionalAssignments_(refDate);

  return Object.freeze(rows.filter(function(item) {
    return item.emailGroups && item.emailGroups.indexOf(groupNorm) >= 0;
  }));
}

function core_getCurrentOccupantsByEmailGroup_(groupName, refDate) {
  var rows = core_getCurrentAssignmentsByEmailGroup_(groupName, refDate);

  var grouped = {};
  rows.forEach(function(item) {
    var key = core_roleProjectionBuildMemberIdentityKey_(item);
    if (!grouped[key]) {
      grouped[key] = {
        identityKey: key,
        memberName: item.memberName,
        email: item.email,
        rga: item.rga,
        roles: []
      };
    }
    grouped[key].roles.push(item);
  });

  Object.keys(grouped).forEach(function(k) {
    grouped[k].roles.sort(function(a, b) {
      return a.displayOrder - b.displayOrder;
    });
    grouped[k] = Object.freeze({
      identityKey: grouped[k].identityKey,
      memberName: grouped[k].memberName,
      email: grouped[k].email,
      rga: grouped[k].rga,
      roles: Object.freeze(grouped[k].roles.slice())
    });
  });

  return Object.freeze(Object.keys(grouped).sort().map(function(k) { return grouped[k]; }));
}

function core_findCurrentMemberContactByIdentity_(person) {
  var sh = core_getSheetByKey_('MEMBERS_ATUAIS');
  if (!sh) return null;

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  var headerMap = core_headerMap_(sh, 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(value) {
    return String(value || '').trim();
  });
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var idxName = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.memberName);
  var idxRga = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.rga);
  var idxEmail = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.email);
  var idxPhone = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.phone);

  idxName = idxName >= 0 ? idxName : -1;
  idxRga = idxRga >= 0 ? idxRga : -1;
  idxEmail = idxEmail >= 0 ? idxEmail : -1;
  idxPhone = idxPhone >= 0 ? idxPhone : -1;

  var targetRga = core_roleProjectionNormalizeKey_(person.rga || '');
  var targetEmail = core_roleProjectionNormalizeKey_(person.email || '');
  var targetName = core_roleProjectionNormalizeKey_(person.memberName || '');

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowRga = idxRga >= 0 ? core_roleProjectionNormalizeKey_(row[idxRga]) : '';
    var rowEmail = idxEmail >= 0 ? core_roleProjectionNormalizeKey_(row[idxEmail]) : '';
    var rowName = idxName >= 0 ? core_roleProjectionNormalizeKey_(row[idxName]) : '';

    var match =
      (targetRga && rowRga && targetRga === rowRga) ||
      (targetEmail && rowEmail && targetEmail === rowEmail) ||
      (targetName && rowName && targetName === rowName);

    if (!match) continue;

    return {
      name: idxName >= 0 ? String(row[idxName] || '').trim() : '',
      rga: idxRga >= 0 ? String(row[idxRga] || '').trim() : '',
      email: idxEmail >= 0 ? String(row[idxEmail] || '').trim() : '',
      phone: idxPhone >= 0 ? String(row[idxPhone] || '').trim() : ''
    };
  }

  return null;
}

function core_formatWhatsappLink_(phone) {
  var raw = String(phone || '').trim();
  if (!raw) return '';

  var digits = raw.replace(/\D/g, '');
  if (!digits) return '';

  if (digits.indexOf('55') !== 0) digits = '55' + digits;
  return 'https://wa.me/' + digits;
}

function core_buildPreferredContactHtml_(person) {
  const phoneRaw =
    String(person.contatoPreferencial || person.telefone || person.whatsapp || '').trim();

  const email =
    String(person.email || '').trim();

  const digits = phoneRaw.replace(/\D+/g, '');

  if (digits.length >= 10) {
    const br = digits.startsWith('55') ? digits : '55' + digits;
    return `<a href="https://wa.me/${br}">${phoneRaw}</a>`;
  }

  if (email) {
    return `<a href="mailto:${email}">${email}</a>`;
  }

  return 'Contato não disponível';
}

function core_getCurrentEmailsByEmailGroup_(groupName, refDate) {
  var people = core_getCurrentOccupantsByEmailGroup_(groupName, refDate);
  return people
    .map(function(p) { return String(p.email || '').trim(); })
    .filter(Boolean);
}

function core_getCurrentEmailsByOccupation_(occupationName, refDate) {
  var assignments = core_getCurrentInstitutionalAssignments_(refDate);
  var emails = assignments
    .filter(function(item) {
      return core_rolesNormalizeText_(item.publicName) === core_rolesNormalizeText_(occupationName);
    })
    .map(function(item) {
      return String(item.email || '').trim();
    })
    .filter(Boolean);

  return Array.from(new Set(emails));
}

function core_getCurrentEmailsByRole_(roleName, refDate) {
  return core_getCurrentEmailsByOccupation_(roleName, refDate);
}

function core_syncMembersCurrentInstitutionalRoles_(refDate) {
  var sh = core_getSheetByKey_(CORE_ROLE_PROJECTION_CURRENT_MEMBERS_KEY);
  if (!sh) throw new Error('Sheet não encontrada para key: ' + CORE_ROLE_PROJECTION_CURRENT_MEMBERS_KEY);

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { updated: 0, total: 0 };

  var headerMap = core_headerMap_(sh, 1);
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var idxName = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.memberName);
  var idxEmail = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.email);
  var idxRga = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.rga);
  var idxCurrentOccupation = core_getPreferredOccupationColumnIndexFromHeaderMap_(headerMap, 'currentOccupation');
  var idxStatus = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.status);

  if (idxCurrentOccupation < 0) {
    throw new Error('Cabecalho de ocupacao atual nao encontrado em MEMBERS_ATUAIS.');
  }

  var grouped = core_groupCurrentInstitutionalAssignmentsByMember_(refDate);
  var updated = 0;

  values.forEach(function(row, i) {
    var memberName = idxName >= 0 ? String(row[idxName] || '').trim() : '';
    var email = idxEmail >= 0 ? String(row[idxEmail] || '').trim() : '';
    var rga = idxRga >= 0 ? String(row[idxRga] || '').trim() : '';
    var status = idxStatus >= 0 ? String(row[idxStatus] || '').trim() : '';

    var identityKey = '';
    if (rga) identityKey = 'RGA::' + core_roleProjectionNormalizeKey_(rga);
    else if (email) identityKey = 'EMAIL::' + core_roleProjectionNormalizeKey_(email);
    else identityKey = 'NOME::' + core_roleProjectionNormalizeKey_(memberName);

    var memberAssignments = grouped[identityKey];
    var newValue = core_buildCurrentOccupationCellValue_(memberAssignments ? memberAssignments.roles : []);

    if (status && core_roleProjectionNormalizeKey_(status) !== 'ATIVO') {
      var currentCell = core_getOccupationValueFromRowByHeaders_(
        row,
        headers,
        'currentOccupation',
        core_roleProjectionNormalizeKey_
      );
      newValue = currentCell.found ? currentCell.value || newValue : newValue;
    }

    if (String(row[idxCurrentOccupation] || '').trim() !== newValue) {
      core_writeOccupationValueByHeaderMap_(sh, i + 2, headerMap, 'currentOccupation', newValue);
      updated++;
    }
  });

  return {
    updated: updated,
    total: values.length
  };
}

function core_syncMembersCurrentDerivedFields_(refDate) {
  var rolesResult = core_syncMembersCurrentInstitutionalRoles_(refDate);

  const sh = core_getSheetByKey_('MEMBERS_ATUAIS');
  if (!sh) {
    return {
      ok: true,
      roles: rolesResult,
      academic: { updatedRows: 0, changedCells: 0 }
    };
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) {
    return {
      ok: true,
      roles: rolesResult,
      academic: { updatedRows: 0, changedCells: 0 }
    };
  }

  const headerMap = core_headerMap_(sh, 1);
  const rgaCol = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.rga);
  const entryCol = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.entrySemester);
  const currentSemesterCol = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.currentSemester);
  const groupSemestersCol = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.groupSemesters);
  const integratedAtCol = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.integratedAt);
  const effectiveGroupTimeCol = core_roleProjectionGuessColumnIndex_(headerMap, CORE_ROLE_PROJECTION_HEADERS.effectiveGroupTime);

  if (rgaCol < 0) throw new Error('core_syncMembersCurrentDerivedFields_: cabecalho "RGA" nao encontrado em MEMBERS_ATUAIS.');
  if (entryCol < 0) throw new Error('core_syncMembersCurrentDerivedFields_: cabecalho "Semestre de Entrada" nao encontrado em MEMBERS_ATUAIS.');
  if (currentSemesterCol < 0) throw new Error('core_syncMembersCurrentDerivedFields_: cabecalho "Semestre atual" nao encontrado em MEMBERS_ATUAIS.');
  if (groupSemestersCol < 0) throw new Error('core_syncMembersCurrentDerivedFields_: cabecalho "QTD_SEMESTRES_NO_GRUPO" nao encontrado em MEMBERS_ATUAIS.');
  if (effectiveGroupTimeCol >= 0 && integratedAtCol < 0) {
    throw new Error('core_syncMembersCurrentDerivedFields_: cabecalho "DATA_INTEGRACAO" nao encontrado em MEMBERS_ATUAIS.');
  }

  const rgaIdx = rgaCol;
  const entryIdx = entryCol;
  const currentSemesterIdx = currentSemesterCol;
  const groupSemestersIdx = groupSemestersCol;
  const integratedAtIdx = integratedAtCol >= 0 ? integratedAtCol : -1;
  const effectiveGroupTimeIdx = effectiveGroupTimeCol >= 0 ? effectiveGroupTimeCol : -1;

  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let changedCells = 0;
  let updatedRows = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    let rowChanged = false;

    const rga = String(row[rgaIdx] || '').trim();
    const entrySemester = String(row[entryIdx] || '').trim();
    const integratedAt = integratedAtIdx >= 0 ? row[integratedAtIdx] : null;

    const currentSemester = rga ? core_getStudentCurrentSemesterFromRga_(rga, refDate) : null;
    const completedGroupSemesterCount = entrySemester
      ? core_getCompletedGroupSemesterCountFromEntrySemester_(entrySemester, refDate)
      : null;
    const civilGroupTime = integratedAt ? core_getCivilGroupTimeFromIntegrationDate_(integratedAt, refDate) : null;

    const currentSemesterDisplay = currentSemester != null ? currentSemester + 'º semestre' : '';
    const groupSemesterDisplay = completedGroupSemesterCount != null ? completedGroupSemesterCount : '0';
    const civilGroupTimeDisplay = integratedAt ? core_formatCivilGroupTime_(civilGroupTime) : '';

    if (String(row[currentSemesterIdx] || '').trim() !== String(currentSemesterDisplay).trim()) {
      row[currentSemesterIdx] = currentSemesterDisplay;
      changedCells++;
      rowChanged = true;
    }

    if (String(row[groupSemestersIdx] || '').trim() !== String(groupSemesterDisplay).trim()) {
      row[groupSemestersIdx] = groupSemesterDisplay;
      changedCells++;
      rowChanged = true;
    }

    if (
      effectiveGroupTimeIdx >= 0 &&
      String(row[effectiveGroupTimeIdx] || '').trim() !== String(civilGroupTimeDisplay).trim()
    ) {
      row[effectiveGroupTimeIdx] = civilGroupTimeDisplay;
      changedCells++;
      rowChanged = true;
    }

    if (rowChanged) updatedRows++;
  }

  if (changedCells > 0) {
    sh.getRange(2, 1, lastRow - 1, lastCol).setValues(values);
  }

  return {
    ok: true,
    roles: rolesResult,
    academic: {
      updatedRows: updatedRows,
      changedCells: changedCells
    }
  };
}

function core_syncMembersCurrentInstitutionalOccupations_(refDate) {
  return core_syncMembersCurrentInstitutionalRoles_(refDate);
}

function core_debugCurrentInstitutionalProjection_(refDate) {
  var assignments = core_getCurrentInstitutionalAssignments_(refDate);
  var grouped = core_groupCurrentInstitutionalAssignmentsByMember_(refDate);

  return {
    totalAssignments: assignments.length,
    totalPeople: Object.keys(grouped).length,
    secretariat: core_getCurrentOccupantsByEmailGroup_('SECRETARIA', refDate)
      .map(function(p) { return p.memberName; }),
    communication: core_getCurrentOccupantsByEmailGroup_('COMUNICACAO', refDate)
      .map(function(p) { return p.memberName; }),
    events: core_getCurrentOccupantsByEmailGroup_('EVENTOS', refDate)
      .map(function(p) { return p.memberName; })
  };
}
