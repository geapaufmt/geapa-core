/***************************************
 * 15_core_roles_projection.js
 *
 * Projeção operacional de cargos atuais
 * a partir de:
 * - VIGENCIA_MEMBROS_DIRETORIAS
 * - VIGENCIA_ASSESSORES
 * - VIGENCIA_CONSELHEIROS
 * + CARGOS_INSTITUCIONAIS_CONFIG
 ***************************************/

const CORE_ROLE_PROJECTION_CURRENT_MEMBERS_KEY = 'MEMBERS_ATUAIS';

const CORE_ROLE_PROJECTION_VIGENCIA_KEYS = Object.freeze({
  diretoria: 'VIGENCIA_MEMBROS_DIRETORIAS',
  assessoria: 'VIGENCIA_ASSESSORES',
  conselho: 'VIGENCIA_CONSELHEIROS'
});

const CORE_ROLE_PROJECTION_HEADERS = Object.freeze({
  memberName: 'MEMBRO',
  memberNameAlt: 'Nome',
  email: 'EMAIL',
  emailAlt: 'E-mail',
  rga: 'RGA',
  currentRole: 'Cargo/função atual',
  status: 'Status'
});

/**
 * ------------------------------------------------------------
 * Normaliza chave de pessoa.
 * ------------------------------------------------------------
 */
function core_roleProjectionNormalizeKey_(value) {
  return String(value == null ? '' : value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

/**
 * ------------------------------------------------------------
 * Retorna data de referência.
 * ------------------------------------------------------------
 */
function core_roleProjectionResolveRefDate_(refDate) {
  return refDate instanceof Date && !isNaN(refDate) ? refDate : new Date();
}

/**
 * ------------------------------------------------------------
 * Cabeçalhos candidatos para nome/email/rga/cargo/data início/fim.
 * ------------------------------------------------------------
 */
function core_roleProjectionGuessColumnIndex_(headerMap, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var col = core_getCol_(headerMap, candidates[i]);
    if (col) return col - 1;
  }
  return -1;
}

/**
 * ------------------------------------------------------------
 * Interpreta se uma linha de vigência está ativa na data ref.
 * Aceita ausência de datas como fallback permissivo.
 * ------------------------------------------------------------
 */
function core_roleProjectionIsAssignmentActive_(startValue, endValue, refDate) {
  var ref = core_startOfDay_(refDate);
  var start = core_roleProjectionParseDate_(startValue);
  var end = core_roleProjectionParseDate_(endValue);

  if (!start && !end) return true;
  if (start && core_startOfDay_(start) > ref) return false;
  if (end && core_startOfDay_(end) < ref) return false;

  return true;
}

/**
 * ------------------------------------------------------------
 * Parse simples de data.
 * ------------------------------------------------------------
 */
function core_roleProjectionParseDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }
  var d = new Date(value);
  return isNaN(d) ? null : d;
}

/**
 * ------------------------------------------------------------
 * Lê assignments ativos de uma sheet de vigência.
 * Compatível com cabeçalhos como:
 * - Nome
 * - RGA
 * - E-mail
 * - Cargo/Função
 * - ID_Diretoria
 * - Data_Início
 * - Data_Fim
 * - Data_Fim_previsto
 * ------------------------------------------------------------
 */
function core_roleProjectionReadAssignmentsFromSheet_(sheetKey, sourceType, refDate) {
  var sh = core_getSheetByKey_(sheetKey);
  if (!sh) {
    throw new Error('Sheet não encontrada para key: ' + sheetKey);
  }

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var headerMap = core_headerMap_(sh, 1);
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var idxName = core_roleProjectionGuessColumnIndex_(headerMap, [
    'Nome',
    'MEMBRO'
  ]);

  var idxEmail = core_roleProjectionGuessColumnIndex_(headerMap, [
    'E-mail',
    'EMAIL'
  ]);

  var idxRga = core_roleProjectionGuessColumnIndex_(headerMap, [
    'RGA'
  ]);

  var idxRole = core_roleProjectionGuessColumnIndex_(headerMap, [
    'Cargo/Função',
    'Cargo',
    'FUNCAO',
    'Função'
  ]);

  var idxStart = core_roleProjectionGuessColumnIndex_(headerMap, [
    'Data_Início',
    'Data início',
    'Data de início',
    'INICIO',
    'Início'
  ]);

  var idxEndReal = core_roleProjectionGuessColumnIndex_(headerMap, [
    'Data_Fim',
    'Data fim',
    'Data de fim',
    'FIM',
    'Fim'
  ]);

  var idxEndPlanned = core_roleProjectionGuessColumnIndex_(headerMap, [
    'Data_Fim_previsto',
    'Data fim previsto',
    'Data de fim prevista',
    'Fim previsto'
  ]);

  var idxAtivo = core_roleProjectionGuessColumnIndex_(headerMap, [
    'ATIVO',
    'Ativo'
  ]);

  var out = [];

  values.forEach(function(row) {
    var memberName = idxName >= 0 ? String(row[idxName] || '').trim() : '';
    var email = idxEmail >= 0 ? String(row[idxEmail] || '').trim() : '';
    var rga = idxRga >= 0 ? String(row[idxRga] || '').trim() : '';
    var rawRole = idxRole >= 0 ? String(row[idxRole] || '').trim() : '';

    if (!memberName && !email && !rga) return;
    if (!rawRole) return;

    // Se existir coluna ATIVO, respeita.
    // Se não existir, considera potencialmente ativo e decide pelas datas.
    if (idxAtivo >= 0) {
      var ativo = core_rolesParseYesNo_(row[idxAtivo]);
      if (!ativo) return;
    }

    var startValue = idxStart >= 0 ? row[idxStart] : null;
    var endRealValue = idxEndReal >= 0 ? row[idxEndReal] : null;
    var endPlannedValue = idxEndPlanned >= 0 ? row[idxEndPlanned] : null;

    // Regra correta:
    // - se há Data_Fim, ela manda
    // - se não há, usa Data_Fim_previsto
    var effectiveEndValue = endRealValue || endPlannedValue;

    if (!core_roleProjectionIsAssignmentActive_(startValue, effectiveEndValue, refDate)) {
      return;
    }

    var roleConfig = core_findInstitutionalRoleByAnyName_(rawRole);

    out.push(Object.freeze({
      sourceType: sourceType,
      rawRole: rawRole,
      roleConfig: roleConfig,
      roleKey: roleConfig ? roleConfig.roleKey : null,
      publicName: roleConfig ? roleConfig.publicName : rawRole,
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

/**
 * ------------------------------------------------------------
 * Lê todos os assignments ativos.
 * ------------------------------------------------------------
 */
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

/**
 * ------------------------------------------------------------
 * Cria chave estável por membro.
 * Preferência: RGA > EMAIL > NOME
 * ------------------------------------------------------------
 */
function core_roleProjectionBuildMemberIdentityKey_(assignment) {
  if (assignment.rgaNorm) return 'RGA::' + assignment.rgaNorm;
  if (assignment.emailNorm) return 'EMAIL::' + assignment.emailNorm;
  return 'NOME::' + assignment.memberNameNorm;
}

/**
 * ------------------------------------------------------------
 * Agrupa assignments por membro.
 * ------------------------------------------------------------
 */
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

/**
 * ------------------------------------------------------------
 * Monta o texto final da célula Cargo/função atual.
 * ------------------------------------------------------------
 */
function core_buildCurrentRoleCellValue_(roles) {
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

/**
 * ------------------------------------------------------------
 * Busca assignment(s) atuais por grupo de e-mail.
 * ------------------------------------------------------------
 */
function core_getCurrentAssignmentsByEmailGroup_(groupName, refDate) {
  var groupNorm = core_rolesNormalizeText_(groupName);
  var rows = core_getCurrentInstitutionalAssignments_(refDate);

  return Object.freeze(rows.filter(function(item) {
    return item.emailGroups && item.emailGroups.indexOf(groupNorm) >= 0;
  }));
}

/**
 * ------------------------------------------------------------
 * Busca ocupantes atuais por grupo de e-mail, já agrupados por membro.
 * ------------------------------------------------------------
 */
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

  return Object.freeze(
    Object.keys(grouped).sort().map(function(k) { return grouped[k]; })
  );
}

/**
 * ------------------------------------------------------------
 * Monta HTML simples de contatos por grupo de e-mail.
 * ------------------------------------------------------------
 */
function core_getCurrentContactsHtmlByEmailGroup_(groupName, refDate) {
  var people = core_getCurrentOccupantsByEmailGroup_(groupName, refDate);
  if (!people.length) {
    return '<p>Nenhum contato disponível no momento.</p>';
  }

  var items = people.map(function(p) {
    var member = core_findCurrentMemberContactByIdentity_(p);

    var displayName = member && member.name ? member.name : p.memberName;
    var phone = member && member.phone ? member.phone : '';
    var email = member && member.email ? member.email : p.email;

    var parts = [];

    if (phone) {
      var wa = core_formatWhatsappLink_(phone);
      parts.push('<a href="' + wa + '">' + phone + '</a>');
    } else if (email) {
      parts.push('<a href="mailto:' + email + '">' + email + '</a>');
    }

    return '<li><strong>' + displayName + '</strong>' +
      (parts.length ? ' — ' + parts.join(' | ') : '') +
      '</li>';
  });

  return '<ul>' + items.join('') + '</ul>';
}

/**
 * ------------------------------------------------------------
 * Sincroniza a coluna Cargo/função atual em MEMBERS_ATUAIS.
 * ------------------------------------------------------------
 */
function core_syncMembersCurrentInstitutionalRoles_(refDate) {
  var sh = core_getSheetByKey_(CORE_ROLE_PROJECTION_CURRENT_MEMBERS_KEY);
  if (!sh) {
    throw new Error('Sheet não encontrada para key: ' + CORE_ROLE_PROJECTION_CURRENT_MEMBERS_KEY);
  }

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { updated: 0, total: 0 };
  }

  var headerMap = core_headerMap_(sh, 1);
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var idxName = core_roleProjectionGuessColumnIndex_(headerMap, [
    CORE_ROLE_PROJECTION_HEADERS.memberName,
    CORE_ROLE_PROJECTION_HEADERS.memberNameAlt
  ]);
  var idxEmail = core_roleProjectionGuessColumnIndex_(headerMap, [
    CORE_ROLE_PROJECTION_HEADERS.email,
    CORE_ROLE_PROJECTION_HEADERS.emailAlt
  ]);
  var idxRga = core_roleProjectionGuessColumnIndex_(headerMap, [
    CORE_ROLE_PROJECTION_HEADERS.rga
  ]);
  var idxCurrentRole = core_roleProjectionGuessColumnIndex_(headerMap, [
    CORE_ROLE_PROJECTION_HEADERS.currentRole
  ]);
  var idxStatus = core_roleProjectionGuessColumnIndex_(headerMap, [
    CORE_ROLE_PROJECTION_HEADERS.status
  ]);

  if (idxCurrentRole < 0) {
    throw new Error('Cabeçalho "Cargo/função atual" não encontrado em MEMBERS_ATUAIS.');
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
    var newValue = core_buildCurrentRoleCellValue_(memberAssignments ? memberAssignments.roles : []);

    // se quiser blindar por status:
    if (status && core_roleProjectionNormalizeKey_(status) !== 'ATIVO') {
      newValue = row[idxCurrentRole] || newValue;
    }

    if (String(row[idxCurrentRole] || '').trim() !== newValue) {
      sh.getRange(i + 2, idxCurrentRole + 1).setValue(newValue);
      updated++;
    }
  });

  return {
    updated: updated,
    total: values.length
  };
}

/**
 * ------------------------------------------------------------
 * Debug geral da projeção atual.
 * ------------------------------------------------------------
 */
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

function core_findCurrentMemberContactByIdentity_(person) {
  var sh = core_getSheetByKey_('MEMBERS_ATUAIS');
  if (!sh) return null;

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  var headerMap = core_headerMap_(sh, 1);
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var idxName = core_getCol_(headerMap, 'MEMBRO') || core_getCol_(headerMap, 'Nome');
  var idxRga = core_getCol_(headerMap, 'RGA');
  var idxEmail = core_getCol_(headerMap, 'EMAIL') || core_getCol_(headerMap, 'E-mail');
  var idxPhone = core_getCol_(headerMap, 'TELEFONE') || core_getCol_(headerMap, 'Telefone');

  idxName = idxName ? idxName - 1 : -1;
  idxRga = idxRga ? idxRga - 1 : -1;
  idxEmail = idxEmail ? idxEmail - 1 : -1;
  idxPhone = idxPhone ? idxPhone - 1 : -1;

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

  // se já vier com 55, mantém; se não, assume Brasil
  if (digits.indexOf('55') !== 0) {
    digits = '55' + digits;
  }

  return 'https://wa.me/' + digits;
}

function core_getCurrentEmailsByEmailGroup_(groupName, refDate) {
  var people = core_getCurrentOccupantsByEmailGroup_(groupName, refDate);
  return people
    .map(function(p) { return String(p.email || '').trim(); })
    .filter(Boolean);
}

function core_getCurrentEmailsByRole_(roleName, refDate) {
  var assignments = core_getCurrentInstitutionalAssignments_(refDate);

  var emails = assignments
    .filter(function(item) {
      return core_rolesNormalizeText_(item.publicName) === core_rolesNormalizeText_(roleName);
    })
    .map(function(item) {
      return String(item.email || '').trim();
    })
    .filter(Boolean);

  return Array.from(new Set(emails));
}