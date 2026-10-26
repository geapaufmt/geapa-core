/***************************************
 * 25_core_ex_members_recipients.js
 *
 * API oficial de destinatarios de ex-membros com consentimento
 * ativo para comunicacoes abertas do GEAPA.
 ***************************************/

const CORE_EX_MEMBERS_COMMUNICATION_CFG = Object.freeze({
  registryKeys: Object.freeze([
    'PESSOAS_EX_MEMBROS',
    'EX_MEMBROS',
    'PESSOAS_EX_MEMBERS',
    'PESSOAS_EX_MEMBROS_BASE'
  ]),
  sheetNames: Object.freeze([
    'Ex-Membros',
    'Ex Membros',
    'Ex_Membros'
  ]),
  headerRow: 1,
  validStatusRegistro: Object.freeze([
    'ATIVO',
    'ATIVA',
    'OK',
    'VALIDO',
    'VALIDA',
    'REGULAR',
    'HOMOLOGADO'
  ]),
  yesValues: Object.freeze(['SIM', 'S', 'TRUE', '1', 'YES']),
  noValues: Object.freeze(['NAO', 'N', 'FALSE', '0', 'NO']),
  eixos: Object.freeze([
    'EIXO_I',
    'EIXO_II',
    'EIXO_III',
    'EIXO_IV',
    'EIXO_V',
    'EIXO_VI',
    'EIXO_VII',
    'EIXO_VIII'
  ]),
  headers: Object.freeze({
    nome: Object.freeze(['NOME', 'NOME_COMPLETO', 'NOME_MEMBRO', 'MEMBRO']),
    rga: Object.freeze(['RGA']),
    email: Object.freeze(['EMAIL', 'E-MAIL', 'EMAIL_PRINCIPAL']),
    recebeComunicacoes: Object.freeze(['RECEBE_COMUNICACOES_GEAPA']),
    statusComunicacao: Object.freeze(['STATUS_COMUNICACAO']),
    statusRegistro: Object.freeze(['STATUS_REGISTRO', 'STATUS', 'STATUS_CADASTRAL']),
    eixosInteresse: Object.freeze(['EIXOS_INTERESSE']),
    eixoFlags: Object.freeze({
      EIXO_I: Object.freeze(['INTERESSE_EIXO_I', 'EIXO_I']),
      EIXO_II: Object.freeze(['INTERESSE_EIXO_II', 'EIXO_II']),
      EIXO_III: Object.freeze(['INTERESSE_EIXO_III', 'EIXO_III']),
      EIXO_IV: Object.freeze(['INTERESSE_EIXO_IV', 'EIXO_IV']),
      EIXO_V: Object.freeze(['INTERESSE_EIXO_V', 'EIXO_V']),
      EIXO_VI: Object.freeze(['INTERESSE_EIXO_VI', 'EIXO_VI']),
      EIXO_VII: Object.freeze(['INTERESSE_EIXO_VII', 'EIXO_VII']),
      EIXO_VIII: Object.freeze(['INTERESSE_EIXO_VIII', 'EIXO_VIII'])
    })
  })
});

function core_normalizeExMemberRecipientText_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function core_normalizeExMemberRecipientKey_(value) {
  return core_normalizeExMemberRecipientText_(value).replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function core_getExMembersCommunicationSheet_() {
  var keys = CORE_EX_MEMBERS_COMMUNICATION_CFG.registryKeys;

  for (var i = 0; i < keys.length; i++) {
    try {
      return core_getSheetByKey_(keys[i]);
    } catch (err) {}
  }

  var raw = core_getRegistryRaw_();
  var currentEnv = core_getCurrentEnv_();
  var wantedNames = CORE_EX_MEMBERS_COMMUNICATION_CFG.sheetNames.map(core_normalizeExMemberRecipientKey_);
  var found = null;

  Object.keys(raw).some(function(key) {
    var envMap = raw[key] || {};
    var entry = envMap[currentEnv] || envMap.PROD || null;
    if (!entry || !entry.ativo) return false;

    var sheetName = core_normalizeExMemberRecipientKey_(entry.sheet);
    if (wantedNames.indexOf(sheetName) === -1) return false;

    found = entry;
    return true;
  });

  if (!found) {
    throw new Error(
      'Nao foi possivel resolver a aba oficial de Ex-Membros. ' +
      'Cadastre no Registry uma KEY oficial como PESSOAS_EX_MEMBROS apontando para a aba "Ex-Membros" da planilha PESSOAS.'
    );
  }

  return core_getSheetById_(found.id, found.sheet);
}

function core_findExMemberRecipientValue_(record, aliases) {
  var names = Array.isArray(aliases) ? aliases : [aliases];
  var keys = Object.keys(record || {});

  for (var i = 0; i < names.length; i++) {
    var wanted = core_normalizeExMemberRecipientKey_(names[i]);
    for (var j = 0; j < keys.length; j++) {
      if (core_normalizeExMemberRecipientKey_(keys[j]) === wanted) {
        return record[keys[j]];
      }
    }
  }

  return '';
}

function core_exMemberRecipientIsYes_(value) {
  return CORE_EX_MEMBERS_COMMUNICATION_CFG.yesValues.indexOf(core_normalizeExMemberRecipientKey_(value)) >= 0;
}

function core_exMemberRecipientIsExplicitNo_(value) {
  return CORE_EX_MEMBERS_COMMUNICATION_CFG.noValues.indexOf(core_normalizeExMemberRecipientKey_(value)) >= 0;
}

function core_exMemberRecipientHasValidStatusRegistro_(value) {
  var status = core_normalizeExMemberRecipientKey_(value);
  if (!status) return false;
  return CORE_EX_MEMBERS_COMMUNICATION_CFG.validStatusRegistro.indexOf(status) >= 0;
}

function core_normalizeExMemberEixo_(value) {
  var raw = core_normalizeExMemberRecipientKey_(value);
  if (!raw) return '';

  var roman = raw
    .replace(/^INTERESSE_EIXO_/, '')
    .replace(/^EIXO_/, '')
    .replace(/^EIXO/, '');

  if (/^(I|II|III|IV|V|VI|VII|VIII)$/.test(roman)) {
    return 'EIXO_' + roman;
  }

  if (CORE_EX_MEMBERS_COMMUNICATION_CFG.eixos.indexOf(raw) >= 0) {
    return raw;
  }

  return raw;
}

function core_parseExMemberEixosList_(value) {
  var text = String(value || '').trim();
  if (!text) return [];

  return text
    .split(/[\r\n,;|]+/)
    .map(core_normalizeExMemberEixo_)
    .filter(Boolean)
    .filter(function(item, index, arr) {
      return arr.indexOf(item) === index;
    });
}

function core_getExMemberEixosFromFlags_(record) {
  var out = [];
  var flags = CORE_EX_MEMBERS_COMMUNICATION_CFG.headers.eixoFlags;

  CORE_EX_MEMBERS_COMMUNICATION_CFG.eixos.forEach(function(eixoKey) {
    var value = core_findExMemberRecipientValue_(record, flags[eixoKey] || []);
    if (core_exMemberRecipientIsYes_(value)) {
      out.push(eixoKey);
    } else if (String(value || '').trim() && !core_exMemberRecipientIsExplicitNo_(value)) {
      var normalized = core_normalizeExMemberEixo_(value);
      if (normalized) out.push(normalized);
    }
  });

  return out.filter(function(item, index, arr) {
    return arr.indexOf(item) === index;
  });
}

function core_getExMemberEixosInteresse_(record) {
  var fromFlags = core_getExMemberEixosFromFlags_(record);
  var fromList = core_parseExMemberEixosList_(
    core_findExMemberRecipientValue_(record, CORE_EX_MEMBERS_COMMUNICATION_CFG.headers.eixosInteresse)
  );

  return fromFlags.concat(fromList).filter(function(item, index, arr) {
    return item && arr.indexOf(item) === index;
  });
}

function core_exMemberEixosIntersect_(candidateEixos, requestedEixos) {
  if (!requestedEixos || !requestedEixos.length) return true;
  var set = {};
  (candidateEixos || []).forEach(function(eixo) {
    set[core_normalizeExMemberEixo_(eixo)] = true;
  });

  return requestedEixos.some(function(eixo) {
    return !!set[core_normalizeExMemberEixo_(eixo)];
  });
}

function core_mapExMemberCommunicationRecipient_(record) {
  var headers = CORE_EX_MEMBERS_COMMUNICATION_CFG.headers;
  var email = core_normalizeEmail_(core_findExMemberRecipientValue_(record, headers.email));

  return Object.freeze({
    nome: String(core_findExMemberRecipientValue_(record, headers.nome) || '').trim(),
    rga: String(core_findExMemberRecipientValue_(record, headers.rga) || '').trim(),
    email: email,
    eixosInteresse: Object.freeze(core_getExMemberEixosInteresse_(record)),
    origem: 'EX_MEMBROS'
  });
}

function core_isEligibleExMemberCommunicationRecord_(record) {
  var headers = CORE_EX_MEMBERS_COMMUNICATION_CFG.headers;
  var email = core_normalizeEmail_(core_findExMemberRecipientValue_(record, headers.email));
  if (!email || !core_isValidEmail_(email)) return false;

  if (!core_exMemberRecipientIsYes_(
    core_findExMemberRecipientValue_(record, headers.recebeComunicacoes)
  )) {
    return false;
  }

  if (core_normalizeExMemberRecipientKey_(
    core_findExMemberRecipientValue_(record, headers.statusComunicacao)
  ) !== 'ATIVO') {
    return false;
  }

  if (!core_exMemberRecipientHasValidStatusRegistro_(
    core_findExMemberRecipientValue_(record, headers.statusRegistro)
  )) {
    return false;
  }

  return true;
}

function core_getExMembersCommunicationRecipients_(options) {
  options = options || {};

  var requestedEixos = core_parseExMemberEixosList_(Array.isArray(options.eixos)
    ? options.eixos.join(',')
    : options.eixos || '');
  var sheet = core_getExMembersCommunicationSheet_();
  var records = core_readSheetRecords_(sheet, {
    headerRow: Number(options.headerRow || CORE_EX_MEMBERS_COMMUNICATION_CFG.headerRow)
  });
  var seenEmails = {};
  var out = [];

  records.forEach(function(record) {
    if (!core_isEligibleExMemberCommunicationRecord_(record)) return;

    var recipient = core_mapExMemberCommunicationRecipient_(record);
    if (!core_exMemberEixosIntersect_(recipient.eixosInteresse, requestedEixos)) return;
    if (seenEmails[recipient.email]) return;

    seenEmails[recipient.email] = true;
    out.push(recipient);
  });

  return Object.freeze(out);
}

function core_debugExMembersCommunicationRecipients_(options) {
  var recipients = core_getExMembersCommunicationRecipients_(options || {});
  return Object.freeze({
    total: recipients.length,
    sample: Object.freeze(recipients.slice(0, 10))
  });
}
