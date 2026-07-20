/** Evolucao cadastral V2: localidades, cursos, dados academicos e setup DEV. */

var CORE_MEMBER_EVOLUTION_SETUP_CONFIRMATION = 'PREPARAR_EVOLUCAO_MEMBROS_V2_DEV';
var CORE_MEMBER_REGISTRATION_PERMISSION = 'membros:cadastrar_novos_membros';

var CORE_MEMBER_EVOLUTION_REGISTRY = Object.freeze({
  CURSOS_CATALOGO: 'PESSOAS_V2_CURSOS_CATALOGO',
  INGRESSOS_MEMBROS: 'PESSOAS_V2_INGRESSOS_MEMBROS',
  CONVITES_AVALIACAO_EGRESSOS: 'PESSOAS_V2_CONVITES_AVALIACAO_EGRESSOS',
  RESPOSTAS_AVALIACAO_EGRESSOS: 'PESSOAS_V2_RESPOSTAS_AVALIACAO_EGRESSOS'
});

var CORE_MEMBER_EVOLUTION_INITIAL_COURSES = Object.freeze([
  Object.freeze({
    CURSO_ID: 'AGRONOMIA_UFMT_SINOP', NOME_CURSO: 'Agronomia', INSTITUICAO: 'UFMT',
    CAMPUS: 'Sinop', NIVEL: 'GRADUACAO', AREA_GERAL: 'CIENCIAS_AGRARIAS', ATIVO: 'SIM',
    ORDEM_EXIBICAO: 10, PERMITE_CADASTRO: 'SIM', OBSERVACOES: 'Linha inicial recomendada; validar institucionalmente antes do setup real.'
  }),
  Object.freeze({
    CURSO_ID: 'OUTRO', NOME_CURSO: 'Outro', INSTITUICAO: '', CAMPUS: '', NIVEL: '',
    AREA_GERAL: '', ATIVO: 'SIM', ORDEM_EXIBICAO: 999, PERMITE_CADASTRO: 'SIM',
    OBSERVACOES: 'Exige nome complementar no contrato de cadastro.'
  })
]);

function coreMemberEvolutionToken_(value) {
  return String(value == null ? '' : value).trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function coreMemberEvolutionText_(value, maxLength, code) {
  var text = String(value == null ? '' : value).trim().replace(/[ \t]+/g, ' ');
  if (text.length > maxLength) {
    var error = new Error(code || 'TEXTO_INVALIDO'); error.code = code || 'TEXTO_INVALIDO'; throw error;
  }
  return text;
}

function coreMemberEvolutionError_(code, message, details) {
  var error = new Error(message || code); error.code = code; error.details = details || {}; return error;
}

function core_getLocalityCatalogV2_(options) {
  var opts = options || {};
  var uf = coreMemberEvolutionToken_(opts.uf);
  var municipalities = uf && CORE_LOCALIDADES_MUNICIPIOS_POR_UF[uf]
    ? CORE_LOCALIDADES_MUNICIPIOS_POR_UF[uf].map(function(item) { return { codigo: String(item[0]), nome: String(item[1]), uf: uf }; })
    : [];
  return Object.freeze({
    metadata: CORE_LOCALIDADES_CATALOGO_META,
    countries: CORE_LOCALIDADES_PAISES,
    states: CORE_LOCALIDADES_UFS,
    municipalities: Object.freeze(municipalities)
  });
}

function core_findCountryV2_(code) {
  var target = coreMemberEvolutionToken_(code);
  for (var i = 0; i < CORE_LOCALIDADES_PAISES.length; i++) {
    if (CORE_LOCALIDADES_PAISES[i].codigo === target) return CORE_LOCALIDADES_PAISES[i];
  }
  return null;
}

function core_findMunicipalityV2_(uf, code) {
  var state = coreMemberEvolutionToken_(uf);
  var target = String(code == null ? '' : code).replace(/\D/g, '');
  var list = CORE_LOCALIDADES_MUNICIPIOS_POR_UF[state] || [];
  for (var i = 0; i < list.length; i++) {
    if (String(list[i][0]) === target) return { codigo: target, nome: String(list[i][1]), uf: state };
  }
  return null;
}

function core_validateOriginV2_(payload) {
  var source = payload || {};
  var countryCode = coreMemberEvolutionToken_(source.paisOrigemCodigo || source.PAIS_ORIGEM_CODIGO);
  var country = core_findCountryV2_(countryCode);
  if (!country) throw coreMemberEvolutionError_('PAIS_ORIGEM_INVALIDO', 'Selecione um pais do catalogo oficial.');

  var city = coreMemberEvolutionText_(source.cidadeNatal || source.cidadeOrigem || source.CIDADE_NATAL, 120, 'CIDADE_ORIGEM_MUITO_LONGA');
  var region = coreMemberEvolutionText_(source.regiaoOrigem || source.REGIAO_ORIGEM, 120, 'REGIAO_ORIGEM_MUITO_LONGA');
  var uf = coreMemberEvolutionToken_(source.ufOrigem || source.UF_ORIGEM);
  var municipalityCode = String(source.municipioOrigemCodigo || source.MUNICIPIO_ORIGEM_CODIGO || '').replace(/\D/g, '');

  if (country.codigo === 'BR') {
    if (!CORE_LOCALIDADES_UFS.some(function(item) { return item.sigla === uf; })) {
      throw coreMemberEvolutionError_('UF_ORIGEM_INVALIDA', 'Selecione uma UF brasileira valida.');
    }
    var municipality = core_findMunicipalityV2_(uf, municipalityCode);
    if (!municipality) throw coreMemberEvolutionError_('MUNICIPIO_ORIGEM_INVALIDO', 'O municipio nao pertence a UF informada.');
    if (city && coreMemberEvolutionToken_(city) !== coreMemberEvolutionToken_(municipality.nome)) {
      throw coreMemberEvolutionError_('MUNICIPIO_ORIGEM_NOME_DIVERGENTE', 'O nome do municipio nao corresponde ao codigo oficial.');
    }
    return Object.freeze({
      PAIS_ORIGEM_CODIGO: 'BR', PAIS_ORIGEM_NOME: country.nome,
      UF_ORIGEM: uf, CIDADE_NATAL: municipality.nome,
      MUNICIPIO_ORIGEM_CODIGO: municipality.codigo, REGIAO_ORIGEM: ''
    });
  }

  if (!city) throw coreMemberEvolutionError_('CIDADE_ORIGEM_OBRIGATORIA', 'Informe a cidade de origem.');
  if (uf || municipalityCode) throw coreMemberEvolutionError_('LOCALIDADE_ESTRANGEIRA_UF_NAO_PERMITIDA', 'UF e codigo de municipio devem ficar vazios para outro pais.');
  return Object.freeze({
    PAIS_ORIGEM_CODIGO: country.codigo, PAIS_ORIGEM_NOME: country.nome,
    UF_ORIGEM: '', CIDADE_NATAL: city, MUNICIPIO_ORIGEM_CODIGO: '', REGIAO_ORIGEM: region
  });
}

function core_validateCourseV2_(payload, courseRecords) {
  var source = payload || {};
  var id = coreMemberEvolutionToken_(source.cursoId || source.CURSO_ID);
  var rows = Array.isArray(courseRecords) ? courseRecords : [];
  var matches = rows.filter(function(row) { return coreMemberEvolutionToken_(row.CURSO_ID) === id; });
  if (matches.length !== 1) throw coreMemberEvolutionError_('CURSO_ID_INVALIDO', 'Curso inexistente ou ambiguo no catalogo.');
  var course = matches[0];
  if (coreMemberEvolutionToken_(course.ATIVO) !== 'SIM' || coreMemberEvolutionToken_(course.PERMITE_CADASTRO) !== 'SIM') {
    throw coreMemberEvolutionError_('CURSO_INATIVO', 'O curso nao esta disponivel para cadastro.');
  }
  var complement = coreMemberEvolutionText_(source.cursoNomeOutro || source.CURSO_NOME_COMPLEMENTAR, 160, 'CURSO_OUTRO_NOME_INVALIDO');
  if (id === 'OUTRO' && !complement) throw coreMemberEvolutionError_('CURSO_OUTRO_EXIGE_DESCRICAO', 'Informe o nome do curso.');
  return Object.freeze({
    CURSO_ID: id,
    CURSO_NOME_SNAPSHOT: id === 'OUTRO' ? complement : String(course.NOME_CURSO || '').trim(),
    INSTITUICAO_ENSINO: id === 'OUTRO' ? coreMemberEvolutionText_(source.instituicaoEnsino || source.INSTITUICAO_ENSINO, 160, 'INSTITUICAO_INVALIDA') : String(course.INSTITUICAO || '').trim(),
    CAMPUS: id === 'OUTRO' ? coreMemberEvolutionText_(source.campus || source.CAMPUS, 120, 'CAMPUS_INVALIDO') : String(course.CAMPUS || '').trim(),
    NIVEL_CURSO: id === 'OUTRO' ? coreMemberEvolutionToken_(source.nivelCurso || source.NIVEL_CURSO) : coreMemberEvolutionToken_(course.NIVEL)
  });
}

function core_parseCivilDateV2_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return { year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate() };
  }
  var match = String(value || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;
  return { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) };
}

function core_civilNumberV2_(value) {
  var parts = core_parseCivilDateV2_(value);
  return parts ? parts.year * 10000 + parts.month * 100 + parts.day : null;
}

function core_parseAcademicEntryPeriodV2_(rga) {
  var parsed = typeof core_parseEntrySemesterFromRga_ === 'function'
    ? core_parseEntrySemesterFromRga_(rga)
    : (function() {
      var digits = String(rga || '').replace(/\D/g, '');
      if (digits.length < 5) return null;
      var year = Number(digits.slice(0, 4)); var semester = Number(digits.charAt(4));
      return year >= 1900 && (semester === 1 || semester === 2) ? { year: year, semester: semester } : null;
    })();
  if (!parsed) return null;
  return Object.freeze({ year: parsed.year, semester: parsed.semester, id: parsed.year + '/' + parsed.semester });
}

function core_calculateAcademicSemesterV2_(rga, semesterRecords, refDate) {
  var entry = core_parseAcademicEntryPeriodV2_(rga);
  if (!entry) throw coreMemberEvolutionError_('RGA_PERIODO_INGRESSO_INVALIDO', 'O RGA nao permite derivar o periodo de ingresso.');
  var referenceNumber = core_civilNumberV2_(refDate || new Date());
  var rows = (Array.isArray(semesterRecords) ? semesterRecords : []).map(function(row) {
    var year = Number(row.ANO || String(row.ID_SEMESTRE || '').split('/')[0]);
    var semester = Number(row.SEMESTRE || String(row.ID_SEMESTRE || '').split('/')[1]);
    return {
      id: year + '/' + semester, year: year, semester: semester,
      start: core_civilNumberV2_(row.DATA_INICIO), end: core_civilNumberV2_(row.DATA_FIM),
      status: coreMemberEvolutionToken_(row.STATUS)
    };
  }).filter(function(row) {
    return row.year && (row.semester === 1 || row.semester === 2) && row.start && row.end && ['CANCELADO','INATIVO'].indexOf(row.status) < 0;
  }).sort(function(a, b) { return a.start - b.start; });

  var entryIndex = rows.map(function(row) { return row.id; }).indexOf(entry.id);
  if (entryIndex < 0) throw coreMemberEvolutionError_('SEMESTRE_INGRESSO_NAO_CATALOGADO', 'O periodo derivado do RGA nao existe no calendario institucional.');
  var targetIndex = -1;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].start <= referenceNumber) targetIndex = i;
  }
  if (targetIndex < entryIndex) throw coreMemberEvolutionError_('SEMESTRE_ATUAL_ANTERIOR_AO_INGRESSO', 'O calendario atual e anterior ao ingresso.');
  return Object.freeze({
    periodoIngressoCurso: entry.id,
    semestreAtualCalculado: targetIndex - entryIndex + 1,
    semestreReferencia: rows[targetIndex].id,
    calculadoEm: refDate || new Date()
  });
}

function core_calculateProfileCompletenessV2_(pessoa, detalhes) {
  var p = pessoa || {}; var d = detalhes || {};
  var required = [
    ['NOME_COMPLETO', p.NOME_COMPLETO], ['EMAIL_PRINCIPAL', p.EMAIL_PRINCIPAL],
    ['DATA_NASCIMENTO', p.DATA_NASCIMENTO], ['PAIS_ORIGEM_CODIGO', p.PAIS_ORIGEM_CODIGO],
    ['CIDADE_NATAL', p.CIDADE_NATAL], ['RGA', d.RGA], ['CURSO_ID', d.CURSO_ID],
    ['PERIODO_INGRESSO_CURSO', d.PERIODO_INGRESSO_CURSO]
  ];
  var present = required.filter(function(item) { return String(item[1] == null ? '' : item[1]).trim(); }).length;
  return Object.freeze({
    status: present === required.length ? 'COMPLETO' : (present === 0 ? 'PENDENTE' : 'PARCIAL'),
    camposAusentes: Object.freeze(required.filter(function(item) { return !String(item[1] == null ? '' : item[1]).trim(); }).map(function(item) { return item[0]; }))
  });
}

function coreMemberEvolutionHeaders_(sheet) {
  if (!sheet) return [];
  var lastColumn = Number(sheet.getLastColumn ? sheet.getLastColumn() : 0);
  if (!lastColumn) return [];
  return (sheet.getRange(1, 1, 1, lastColumn).getValues()[0] || []).map(function(item) { return String(item || '').trim(); });
}

function coreMemberEvolutionRecords_(sheet, headers) {
  if (!sheet || !headers.length || Number(sheet.getLastRow()) < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return values.map(function(row) {
    var record = {}; headers.forEach(function(header, index) { record[header] = row[index]; }); return record;
  });
}

function coreMemberEvolutionPlanSheet_(name, expected, sheet) {
  var existing = coreMemberEvolutionHeaders_(sheet);
  var normalized = existing.map(coreMemberEvolutionToken_);
  var missing = expected.filter(function(header) { return normalized.indexOf(coreMemberEvolutionToken_(header)) < 0; });
  return {
    ambiente: 'DEV', planilha: 'PESSOAS_V2_DB', aba: name,
    cabecalhosExistentes: existing, cabecalhosAusentes: missing,
    linhasRecomendadas: [], alteracoesPlanejadas: missing.length ? ['ADICIONAR_CABECALHOS_AO_FINAL'] : [],
    escritaExecutada: false, idempotente: true
  };
}

function coreMemberEvolutionEnsureHeaders_(sheet, expected) {
  var plan = coreMemberEvolutionPlanSheet_(sheet.getName(), expected, sheet);
  if (plan.cabecalhosAusentes.length) {
    var start = Math.max(1, plan.cabecalhosExistentes.length + 1);
    sheet.getRange(1, start, 1, plan.cabecalhosAusentes.length).setValues([plan.cabecalhosAusentes]);
  }
  return plan;
}

function coreMemberEvolutionEnsureRegistry_(spreadsheetId, specs) {
  var registry = core_openSpreadsheetById_(CORE_REGISTRY_SPREADSHEET_ID).getSheetByName(CORE_REGISTRY_SHEET_NAME);
  if (!registry) throw coreMemberEvolutionError_('REGISTRY_SHEET_INDISPONIVEL', 'Registry indisponivel.');
  var data = core_readSheetData_(registry, { headerRow: 1 });
  var headers = data.headers;
  specs.forEach(function(spec) {
    var found = data.rows.map(function(row) { return core_rowToObject_(headers, row); }).filter(function(row) {
      return coreMemberEvolutionToken_(row.KEY) === spec.key && coreMemberEvolutionToken_(row.AMBIENTE) === 'DEV' && coreMemberEvolutionToken_(row.ATIVO) === 'SIM';
    });
    if (found.length > 1) throw coreMemberEvolutionError_('REGISTRY_DEV_DUPLICADO_' + spec.key, 'Registry DEV duplicado.');
    if (found.length === 1) {
      if (String(found[0].SPREADSHEET_ID || '').trim() !== spreadsheetId || String(found[0].SHEET_NAME || '').trim() !== spec.sheetName) {
        throw coreMemberEvolutionError_('REGISTRY_DEV_DIVERGENTE_' + spec.key, 'Registry DEV divergente.');
      }
      return;
    }
    var values = { KEY: spec.key, SPREADSHEET_ID: spreadsheetId, SHEET_NAME: spec.sheetName, ATIVO: 'SIM', AMBIENTE: 'DEV', DISPLAY_NAME: spec.sheetName, TYPE: spec.type, NOTAS: 'Fonte DEV/HOMOLOG da evolucao de membros V2.' };
    registry.appendRow(corePerfilRegistryRowValues_(headers, Object.keys(values).reduce(function(out, key) { out[core_normalizeHeader_(key)] = values[key]; return out; }, {})));
  });
  core_registryCacheClear_();
}

function coreMemberEvolutionEnsurePermission_() {
  var source = corePerfilReadSheetSource_(corePerfilGetSheetByKey_('PORTAL_PERMISSOES', 'DEV'), 'PORTAL_PERMISSOES');
  ['ADMIN','DIRETORIA'].forEach(function(profile) {
    var exists = source.records.some(function(row) {
      return coreMemberEvolutionToken_(row.PERFIL_PORTAL) === profile && String(row.PERMISSAO || '').trim().toLowerCase() === CORE_MEMBER_REGISTRATION_PERMISSION && coreMemberEvolutionToken_(row.ATIVO) === 'SIM';
    });
    if (!exists) corePerfilAppendRecord_(source, { PERFIL_PORTAL: profile, PERMISSAO: CORE_MEMBER_REGISTRATION_PERMISSION, ATIVO: 'SIM', OBS: 'Cadastro de pessoas ja aprovadas pela Diretoria.' });
  });
}

function core_setupMemberEvolutionV2_(options) {
  var opts = options || {};
  var environment = coreMemberEvolutionToken_(opts.ambiente || opts.environment || 'DEV');
  if (environment === 'HOMOLOG') environment = 'DEV';
  if (environment !== 'DEV') throw coreMemberEvolutionError_('SETUP_EVOLUCAO_MEMBROS_PROD_PROIBIDO', 'Este setup aceita somente DEV/HOMOLOG.');
  var dryRun = opts.dryRun !== false;
  var workbook = opts.spreadsheet || core_openDomainSpreadsheet_('PESSOAS', { ambiente: 'DEV', forWrite: !dryRun });
  var spreadsheetId = String(opts.spreadsheetId || (workbook.getId ? workbook.getId() : 'INJETADO_TESTE'));
  var report = {
    ambiente: 'DEV', dryRun: dryRun, tokenConfirmacao: CORE_MEMBER_EVOLUTION_SETUP_CONFIRMATION,
    escritaExecutada: false, idempotente: true, planilha: 'PESSOAS_V2_DB', abas: [],
    registryDev: [], permissoesDev: { permissao: CORE_MEMBER_REGISTRATION_PERMISSION, perfis: ['ADMIN','DIRETORIA'], secretariaAutomatica: false },
    mailHub: { configuracoesRecomendadas: ['PORTAL_HOMOLOG_URL','MEMBROS_EGRESS_FEEDBACK_BASE_URL'], remetenteHardcoded: false }
  };
  Object.keys(CORE_MEMBER_EVOLUTION_SCHEMAS).forEach(function(name) {
    var sheet = workbook.getSheetByName(name);
    var plan = coreMemberEvolutionPlanSheet_(name, CORE_MEMBER_EVOLUTION_SCHEMAS[name], sheet);
    if (name === 'CURSOS_CATALOGO') {
      var records = coreMemberEvolutionRecords_(sheet, plan.cabecalhosExistentes);
      var ids = records.map(function(row) { return coreMemberEvolutionToken_(row.CURSO_ID); });
      plan.linhasRecomendadas = CORE_MEMBER_EVOLUTION_INITIAL_COURSES.filter(function(row) { return ids.indexOf(row.CURSO_ID) < 0; });
      if (plan.linhasRecomendadas.length) plan.alteracoesPlanejadas.push('ADICIONAR_LINHAS_INICIAIS_AUSENTES');
    }
    report.abas.push(plan);
    if (CORE_MEMBER_EVOLUTION_REGISTRY[name]) report.registryDev.push({ key: CORE_MEMBER_EVOLUTION_REGISTRY[name], aba: name, spreadsheetIdMascarado: spreadsheetId.slice(0, 4) + '***' + spreadsheetId.slice(-4) });
  });
  if (dryRun) return Object.freeze(report);
  if (String(opts.confirmacao || '').trim() !== CORE_MEMBER_EVOLUTION_SETUP_CONFIRMATION) {
    throw coreMemberEvolutionError_('CONFIRMACAO_SETUP_INVALIDA', 'Informe o token exato de confirmacao.');
  }
  return core_withLock_('SETUP_EVOLUCAO_MEMBROS_V2_DEV', function() {
    Object.keys(CORE_MEMBER_EVOLUTION_SCHEMAS).forEach(function(name) {
      var sheet = workbook.getSheetByName(name) || workbook.insertSheet(name);
      coreMemberEvolutionEnsureHeaders_(sheet, CORE_MEMBER_EVOLUTION_SCHEMAS[name]);
      if (name === 'CURSOS_CATALOGO') {
        var headers = coreMemberEvolutionHeaders_(sheet);
        var records = coreMemberEvolutionRecords_(sheet, headers);
        var ids = records.map(function(row) { return coreMemberEvolutionToken_(row.CURSO_ID); });
        CORE_MEMBER_EVOLUTION_INITIAL_COURSES.forEach(function(row) {
          if (ids.indexOf(row.CURSO_ID) >= 0) return;
          var now = new Date(); var record = Object.assign({}, row, { CRIADO_EM: now, ATUALIZADO_EM: now });
          sheet.appendRow(core_buildRowFromObjectByHeaders_(headers, record));
        });
      }
    });
    coreMemberEvolutionEnsureRegistry_(spreadsheetId, Object.keys(CORE_MEMBER_EVOLUTION_REGISTRY).map(function(name) {
      return {
        key: CORE_MEMBER_EVOLUTION_REGISTRY[name],
        sheetName: name,
        type: (name === 'CURSOS_CATALOGO' || name === 'RESPOSTAS_AVALIACAO_EGRESSOS')
          ? 'FONTE'
          : 'FONTE_EVENTOS'
      };
    }));
    coreMemberEvolutionEnsurePermission_();
    report.escritaExecutada = true; report.dryRun = false;
    return Object.freeze(report);
  }, 30000);
}
