/**
 * ============================================================
 * 29_core_domains_v2_migration.js
 * ============================================================
 *
 * Arquivo historico da migracao v2. Nao chamar em producao.
 * Mantido temporariamente apenas para referencia. Remover apos
 * validacao do portal v2.
 *
 * Arquivo historico da migracao v2. Nao chamar em producao.
 * Mantido temporariamente apenas para referencia. Remover apos
 * validacao do portal v2.
 *
 * Migracao controlada dos legados para PESSOAS v2 e VIGENCIAS v2.
 *
 * Regras de seguranca:
 * - fontes legadas sao somente leitura;
 * - destino e sempre v2;
 * - dryRun e true por padrao;
 * - resetDestino nao e executado automaticamente;
 * - upserts usam chaves naturais para evitar duplicidade.
 */

var CORE_DOMAINS_V2_MIGRATION_CFG = Object.freeze({
  pessoasTitle: 'PESSOAS v2 - DEV',
  vigenciasTitle: 'VIGENCIAS v2 - DEV',
  defaultOptions: Object.freeze({
    dryRun: true,
    limit: null,
    resetDestino: false,
    escreverRelatorio: true
  }),
  pessoasSources: Object.freeze({
    membrosAtuais: Object.freeze({
      label: 'Membros Atuais',
      registryKeys: Object.freeze(['MEMBERS_ATUAIS', 'PESSOAS_MEMBROS_ATUAIS']),
      sheetNames: Object.freeze(['Membros Atuais', 'MEMBERS_ATUAIS'])
    }),
    exMembros: Object.freeze({
      label: 'Ex-Membros',
      registryKeys: Object.freeze(['PESSOAS_EX_MEMBROS', 'EX_MEMBROS']),
      sheetNames: Object.freeze(['Ex-Membros', 'Ex Membros', 'Ex_Membros'])
    }),
    membrosEspera: Object.freeze({
      label: 'Membros em Espera',
      registryKeys: Object.freeze(['PESSOAS_MEMBROS_EM_ESPERA', 'MEMBROS_EM_ESPERA']),
      sheetNames: Object.freeze(['Membros em Espera', 'Membros Espera', 'MEMBROS_EM_ESPERA'])
    }),
    eventos: Object.freeze({
      label: 'Membros_Eventos',
      registryKeys: Object.freeze(['MEMBER_EVENTOS_VINCULO', 'MEMBROS_EVENTOS_VINCULO', 'PESSOAS_MEMBROS_EVENTOS']),
      sheetNames: Object.freeze(['Membros_Eventos', 'Membros Eventos', 'MEMBROS_EVENTOS_VINCULO'])
    }),
    professores: Object.freeze({
      label: 'Dados dos Professores/Tecnicos',
      registryKeys: Object.freeze(['PESSOAS_PROFESSORES_TECNICOS', 'PROFESSORES_TECNICOS']),
      sheetNames: Object.freeze(['Dados dos Professores/Técnicos', 'Dados dos Professores/Tecnicos', 'Professores/Tecnicos'])
    }),
    externos: Object.freeze({
      label: 'Participantes Externos',
      registryKeys: Object.freeze(['PESSOAS_PARTICIPANTES_EXTERNOS', 'PARTICIPANTES_EXTERNOS']),
      sheetNames: Object.freeze(['Participantes Externos', 'Participantes_Externos'])
    })
  }),
  vigenciasSources: Object.freeze({
    semestres: Object.freeze({
      label: 'Semestres',
      registryKeys: Object.freeze(['VIGENCIA_SEMESTRES']),
      sheetNames: Object.freeze(['Semestres', 'SEMESTRES'])
    }),
    periodos: Object.freeze({
      label: 'Periodos',
      registryKeys: Object.freeze(['VIGENCIA_PERIODOS', 'VIGENCIA_PERIODOS_GEAPA']),
      sheetNames: Object.freeze(['Períodos', 'Periodos', 'PERIODOS'])
    }),
    diretorias: Object.freeze({
      label: 'Diretorias',
      registryKeys: Object.freeze(['VIGENCIA_DIRETORIAS']),
      sheetNames: Object.freeze(['Diretorias', 'DIRETORIAS'])
    }),
    semestresDiretoria: Object.freeze({
      label: 'Semestres_Diretoria',
      registryKeys: Object.freeze(['VIGENCIA_SEMESTRES_DIRETORIA']),
      sheetNames: Object.freeze(['Semestres_Diretoria', 'SEMESTRES_DIRETORIA'])
    }),
    cargosConfig: Object.freeze({
      label: 'Cargos_Config',
      registryKeys: Object.freeze(['VIGENCIA_CARGOS_CONFIG', 'CARGOS_CONFIG']),
      sheetNames: Object.freeze(['Cargos_Config', 'CARGOS_CONFIG'])
    }),
    diretores: Object.freeze({
      label: 'Diretores',
      registryKeys: Object.freeze(['VIGENCIA_MEMBROS_DIRETORIAS']),
      sheetNames: Object.freeze(['Diretores', 'Membros_Diretorias', 'VIGENCIA_MEMBROS_DIRETORIAS'])
    }),
    assessores: Object.freeze({
      label: 'Assessores',
      registryKeys: Object.freeze(['VIGENCIA_ASSESSORES']),
      sheetNames: Object.freeze(['Assessores', 'VIGENCIA_ASSESSORES'])
    }),
    conselheiros: Object.freeze({
      label: 'Conselheiros',
      registryKeys: Object.freeze(['VIGENCIA_CONSELHEIROS']),
      sheetNames: Object.freeze(['Conselheiros', 'VIGENCIA_CONSELHEIROS'])
    })
  })
});

function core_domainsV2MigrationOptions_(options) {
  options = options || {};
  var defaults = CORE_DOMAINS_V2_MIGRATION_CFG.defaultOptions;
  var out = {
    dryRun: options.dryRun !== false,
    limit: options.limit == null || options.limit === '' ? null : Number(options.limit),
    resetDestino: options.resetDestino === true,
    escreverRelatorio: options.escreverRelatorio !== false,
    confirmacao: String(options.confirmacao || '').trim()
  };
  if (out.resetDestino && out.confirmacao !== 'RESETAR_DESTINO_V2') {
    throw new Error('resetDestino exige confirmacao: "RESETAR_DESTINO_V2". Planilhas antigas nunca sao afetadas.');
  }
  if (out.limit != null && (!isFinite(out.limit) || out.limit < 1)) {
    out.limit = defaults.limit;
  }
  return out;
}

function core_domainsV2NormalizeText_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function core_domainsV2NormalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function core_domainsV2OnlyDigits_(value) {
  return String(value || '').replace(/\D+/g, '');
}

function core_domainsV2GetRga_(record) {
  return String(core_domainsV2GetByAliases_(record, [
    'RGA',
    'Registro Geral Academico',
    'REGISTRO_GERAL_ACADEMICO'
  ]) || '').trim();
}

function core_domainsV2GetByAliases_(record, aliases) {
  aliases = aliases || [];
  var keys = Object.keys(record || {});
  var normalized = {};
  keys.forEach(function(key) {
    normalized[core_normalizeHeader_(key)] = key;
  });
  for (var i = 0; i < aliases.length; i++) {
    var wanted = core_normalizeHeader_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(normalized, wanted)) {
      var value = record[normalized[wanted]];
      if (value !== '' && value != null) return value;
    }
  }
  return '';
}

function core_domainsV2ReadRecordsSafe_(sheet) {
  if (!sheet) return [];
  return core_readSheetRecords_(sheet, { skipBlankRows: true });
}

function core_domainsV2FindSheetFromSource_(sourceCfg) {
  var errors = [];
  var registryKeys = sourceCfg.registryKeys || [];
  for (var i = 0; i < registryKeys.length; i++) {
    try {
      return {
        sheet: core_getSheetByKey_(registryKeys[i]),
        source: 'REGISTRY',
        key: registryKeys[i]
      };
    } catch (err) {
      errors.push(registryKeys[i] + ': ' + err.message);
    }
  }

  try {
    var currentEnv = core_getCurrentEnv_();
    var raw = core_getRegistryRaw_();
    var wanted = (sourceCfg.sheetNames || []).map(core_domainsV2NormalizeText_);
    var keys = Object.keys(raw);
    for (var r = 0; r < keys.length; r++) {
      var envMap = raw[keys[r]];
      var envs = Object.keys(envMap || {});
      for (var e = 0; e < envs.length; e++) {
        var entry = envMap[envs[e]];
        if (!entry || !entry.ativo) continue;
        if (!core_registryMatchesEnv_(entry.ambiente, currentEnv)) continue;
        if (wanted.indexOf(core_domainsV2NormalizeText_(entry.sheet)) < 0) continue;
        return {
          sheet: core_getSheetById_(entry.id, entry.sheet),
          source: 'REGISTRY_SHEET_NAME',
          key: entry.key
        };
      }
    }
  } catch (fallbackErr) {
    errors.push('fallback por nome: ' + fallbackErr.message);
  }

  return {
    sheet: null,
    source: 'NOT_FOUND',
    key: '',
    errors: errors
  };
}

function core_domainsV2OpenDestSpreadsheet_(domain) {
  var id = CORE_DOMAINS_V2_DEV_SPREADSHEETS[domain];
  var spreadsheet = SpreadsheetApp.openById(id);
  if (domain === 'PESSOAS') spreadsheet.rename(CORE_DOMAINS_V2_MIGRATION_CFG.pessoasTitle);
  if (domain === 'VIGENCIAS') spreadsheet.rename(CORE_DOMAINS_V2_MIGRATION_CFG.vigenciasTitle);
  try {
    spreadsheet.setSpreadsheetTimeZone('America/Cuiaba');
  } catch (err) {
    // Algumas contas/API podem nao permitir ajuste de timezone; diagnostico registra se necessario.
  }
  return spreadsheet;
}

function core_domainsV2ReadDest_(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error('Aba destino v2 nao encontrada: ' + sheetName);
  return {
    sheet: sheet,
    records: core_readSheetRecords_(sheet, { skipBlankRows: true })
  };
}

function core_domainsV2IndexByField_(records, field) {
  var out = {};
  records.forEach(function(record) {
    var value = String(record[field] || '').trim();
    if (value) out[value] = record;
  });
  return out;
}

function core_domainsV2MaxSeq_(records, field, prefix) {
  var max = 0;
  records.forEach(function(record) {
    var value = String(record[field] || '');
    var match = value.match(new RegExp('^' + prefix + '-(\\d+)$'));
    if (match) max = Math.max(max, Number(match[1]));
  });
  return max;
}

function core_domainsV2SeqId_(prefix, number) {
  var text = String(number);
  while (text.length < 6) text = '0' + text;
  return prefix + '-' + text;
}

function core_domainsV2AppendRows_(sheet, rows, dryRun) {
  if (dryRun || !rows.length) return 0;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(header) { return String(header || '').trim(); });
  var values = rows.map(function(row) {
    return core_buildRowFromObjectByHeaders_(headers, row);
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
  return values.length;
}

function core_domainsV2ClearDataRowsForDomain_(domain, options) {
  options = options || {};
  if (String(options.confirmacao || '').trim() !== 'RESETAR_DESTINO_V2') {
    throw new Error('Reset do destino v2 exige confirmacao: "RESETAR_DESTINO_V2".');
  }

  var ss = core_domainsV2OpenDestSpreadsheet_(domain);
  var schemas = CORE_DOMAINS_V2_SCHEMAS[domain];
  var out = [];
  schemas.forEach(function(definition) {
    var sheet = ss.getSheetByName(definition.sheetName);
    if (!sheet) {
      out.push({
        sheetName: definition.sheetName,
        ok: false,
        error: 'Aba nao encontrada.'
      });
      return;
    }

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow > 1 && lastColumn > 0) {
      sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
    }
    out.push({
      sheetName: definition.sheetName,
      ok: true,
      rowsCleared: Math.max(0, lastRow - 1)
    });
  });

  return {
    domain: domain,
    ok: out.every(function(item) { return item.ok; }),
    sheets: out
  };
}

function core_domainsV2NaturalPersonKey_(kind, record) {
  var rga = core_domainsV2GetRga_(record);
  var idProfessor = String(core_domainsV2GetByAliases_(record, ['ID_PROFESSOR', 'ID Professor']) || '').trim();
  var idExterno = String(core_domainsV2GetByAliases_(record, ['ID_PARTICIPANTE_EXTERNO', 'ID Externo']) || '').trim();
  var email = core_domainsV2NormalizeEmail_(core_domainsV2GetByAliases_(record, ['EMAIL', 'E-mail', 'EMAIL_PRINCIPAL']));
  var cpf = core_domainsV2OnlyDigits_(core_domainsV2GetByAliases_(record, ['CPF']));
  if (kind === 'PROFESSOR' && idProfessor) return 'ID_PROFESSOR:' + idProfessor;
  if (kind === 'EXTERNO' && idExterno) return 'ID_EXTERNO:' + idExterno;
  if (kind === 'EXTERNO' && email) return 'EMAIL:' + email;
  if (rga) return 'RGA:' + rga;
  if (email) return 'EMAIL:' + email;
  if (cpf) return 'CPF:' + cpf;
  return '';
}

function core_domainsV2IdentifierMapKey_(type, value) {
  type = String(type || '').trim().toUpperCase();
  if (type === 'EMAIL') value = core_domainsV2NormalizeEmail_(value);
  else if (type === 'CPF') value = core_domainsV2OnlyDigits_(value);
  else value = String(value || '').trim();
  return type && value ? type + ':' + value : '';
}

function core_domainsV2RecordName_(record) {
  return core_domainsV2NormalizeText_(core_domainsV2GetByAliases_(record, [
    'NOME_COMPLETO',
    'NOME_MEMBRO',
    'NOME_EXIBICAO',
    'NOME',
    'Nome completo',
    'Nome'
  ]));
}

function core_domainsV2BaseCompatibleWithRecord_(baseRecord, incomingRecord) {
  if (!baseRecord) return false;

  var baseEmail = core_domainsV2NormalizeEmail_(baseRecord.EMAIL_PRINCIPAL);
  var incomingEmail = core_domainsV2NormalizeEmail_(core_domainsV2GetByAliases_(incomingRecord, ['EMAIL_PRINCIPAL', 'EMAIL', 'E-mail']));
  if (baseEmail && incomingEmail && baseEmail !== incomingEmail) return false;

  var baseCpf = core_domainsV2OnlyDigits_(baseRecord.CPF);
  var incomingCpf = core_domainsV2OnlyDigits_(core_domainsV2GetByAliases_(incomingRecord, ['CPF']));
  if (baseCpf && incomingCpf && baseCpf !== incomingCpf) return false;

  var baseName = core_domainsV2NormalizeText_(baseRecord.NOME_COMPLETO || baseRecord.NOME_EXIBICAO);
  var incomingName = core_domainsV2RecordName_(incomingRecord);
  if (baseName && incomingName && baseName !== incomingName) {
    var baseParts = baseName.split(' ').filter(function(part) { return part.length >= 3; });
    var incomingParts = incomingName.split(' ').filter(function(part) { return part.length >= 3; });
    var common = baseParts.some(function(part) { return incomingParts.indexOf(part) >= 0; });
    if (!common) return false;
  }

  return true;
}

function core_domainsV2BuildPessoaBase_(idPessoa, kind, record, sourceLabel) {
  var name = core_domainsV2GetByAliases_(record, ['NOME_COMPLETO', 'NOME_MEMBRO', 'NOME', 'Nome completo', 'Nome']);
  var email = core_domainsV2NormalizeEmail_(core_domainsV2GetByAliases_(record, ['EMAIL_PRINCIPAL', 'EMAIL', 'E-mail']));
  return {
    ID_PESSOA: idPessoa,
    NOME_COMPLETO: String(name || '').trim(),
    NOME_EXIBICAO: String(core_domainsV2GetByAliases_(record, ['NOME_EXIBICAO', 'NOME_SOCIAL', 'NOME']) || name || '').trim(),
    EMAIL_PRINCIPAL: email,
    TELEFONE_PRINCIPAL: core_domainsV2GetByAliases_(record, ['TELEFONE_PRINCIPAL', 'TELEFONE', 'Telefone']),
    CPF: core_domainsV2OnlyDigits_(core_domainsV2GetByAliases_(record, ['CPF'])),
    DATA_NASCIMENTO: core_domainsV2GetByAliases_(record, ['DATA_NASCIMENTO', 'Data de nascimento']),
    INSTAGRAM: core_domainsV2GetByAliases_(record, ['INSTAGRAM']),
    CIDADE_NATAL: core_domainsV2GetByAliases_(record, ['CIDADE_NATAL', 'Cidade natal', 'CIDADE']),
    UF_ORIGEM: core_domainsV2GetByAliases_(record, ['UF_ORIGEM', 'UF']),
    SEXO: core_domainsV2GetByAliases_(record, ['SEXO']),
    STATUS_CADASTRAL: core_domainsV2GetByAliases_(record, ['STATUS_CADASTRAL', 'STATUS']) || 'ATIVO',
    OBS_INTERNA: 'Migrado de fonte legada: ' + sourceLabel + ' (' + kind + ').',
    CRIADO_EM: new Date(),
    ATUALIZADO_EM: new Date(),
    ATIVO: 'SIM'
  };
}

function core_domainsV2AddIdentifier_(state, idPessoa, type, value, principal) {
  value = String(value || '').trim();
  if (!value) return;
  var mapKey = core_domainsV2IdentifierMapKey_(type, value);
  if (mapKey && state.personKeyToId[mapKey] && state.personKeyToId[mapKey] !== idPessoa) {
    state.identifierConflicts.push({
      tipo: 'IDENTIFICADOR_JA_ASSOCIADO_A_OUTRA_PESSOA',
      identificador: mapKey,
      idPessoaExistente: state.personKeyToId[mapKey],
      idPessoaNovo: idPessoa
    });
    return;
  }
  var key = idPessoa + '|' + type + '|' + value;
  if (state.identifierKeys[key]) return;
  state.identifierKeys[key] = true;
  if (mapKey && !state.personKeyToId[mapKey]) {
    state.personKeyToId[mapKey] = idPessoa;
  }
  state.nextIdentificador++;
  state.identificadores.push({
    ID_IDENTIFICADOR: core_domainsV2SeqId_('IDE', state.nextIdentificador),
    ID_PESSOA: idPessoa,
    TIPO_IDENTIFICADOR: type,
    VALOR_IDENTIFICADOR: value,
    PRINCIPAL: principal ? 'SIM' : 'NAO',
    ATIVO: 'SIM',
    OBS: 'Migracao v2.'
  });
}

function core_domainsV2AddVinculo_(state, idPessoa, tipo, status, record, sourceLabel) {
  var key = idPessoa + '|' + tipo;
  if (state.vinculoKeys[key]) return;
  state.vinculoKeys[key] = true;
  state.nextVinculo++;
  state.vinculos.push({
    ID_VINCULO: core_domainsV2SeqId_('VIN', state.nextVinculo),
    ID_PESSOA: idPessoa,
    TIPO_VINCULO: tipo,
    STATUS_VINCULO: status,
    DATA_INICIO: core_domainsV2GetByAliases_(record, ['DATA_INTEGRACAO', 'DATA_INICIO', 'Data_Inicio', 'DATA_ENTRADA']),
    DATA_FIM: core_domainsV2GetByAliases_(record, ['DATA_FIM', 'DATA_DESLIGAMENTO', 'DATA_SAIDA']),
    MOTIVO_INICIO: sourceLabel,
    MOTIVO_FIM: core_domainsV2GetByAliases_(record, ['MOTIVO_FIM', 'MOTIVO_DESLIGAMENTO']),
    FONTE: 'MIGRACAO_V2',
    LINK_ATA_OU_PROCESSO: core_domainsV2GetByAliases_(record, ['LINK_ATA', 'LINK_PROCESSO']),
    OBS_PUBLICA: '',
    OBS_INTERNA: 'Vinculo gerado por migracao controlada v2.',
    ATIVO: status === 'ATIVO' || status === 'EM_ESPERA' || status === 'SUSPENSO' ? 'SIM' : 'NAO'
  });
}

function core_domainsV2EnsurePessoa_(state, kind, record, sourceLabel, report) {
  var key = core_domainsV2NaturalPersonKey_(kind, record);
  var name = core_domainsV2GetByAliases_(record, ['NOME_COMPLETO', 'NOME_MEMBRO', 'NOME', 'Nome completo', 'Nome']);
  var email = core_domainsV2NormalizeEmail_(core_domainsV2GetByAliases_(record, ['EMAIL_PRINCIPAL', 'EMAIL', 'E-mail']));
  var cpf = core_domainsV2OnlyDigits_(core_domainsV2GetByAliases_(record, ['CPF']));
  var rga = core_domainsV2GetRga_(record);

  if (!name) report.registrosSemNome++;
  if (!email) report.registrosSemEmail++;
  if ((kind === 'MEMBRO' || kind === 'EGRESSO' || kind === 'EX_MEMBRO' || kind === 'ESPERA') && !rga) report.registrosSemRga++;
  if (!key) {
    report.revisoes.push({ tipo: 'PESSOA_SEM_CHAVE_DEDUP', fonte: sourceLabel, linha: record.__rowNumber || '' });
    return '';
  }
  if (email) report.emailCounts[email] = (report.emailCounts[email] || 0) + 1;
  if (cpf) report.cpfCounts[cpf] = (report.cpfCounts[cpf] || 0) + 1;

  if (state.personKeyToId[key]) {
    var existingId = state.personKeyToId[key];
    if (core_domainsV2BaseCompatibleWithRecord_(state.pessoasById[existingId], record)) {
      return existingId;
    }
    report.revisoes.push({
      tipo: 'ID_PESSOA_EXISTENTE_INCOMPATIVEL_COM_REGISTRO_LEGADO',
      chave: key,
      idPessoaExistente: existingId,
      fonte: sourceLabel,
      linha: record.__rowNumber || ''
    });
  }

  state.nextPessoa++;
  var idPessoa = core_domainsV2SeqId_('PES', state.nextPessoa);
  state.personKeyToId[key] = idPessoa;
  var base = core_domainsV2BuildPessoaBase_(idPessoa, kind, record, sourceLabel);
  state.pessoasById[idPessoa] = base;
  state.pessoas.push(base);

  if ((kind === 'MEMBRO' || kind === 'EGRESSO' || kind === 'EX_MEMBRO' || kind === 'ESPERA') && rga) {
    core_domainsV2AddIdentifier_(state, idPessoa, 'RGA', rga, true);
  }
  if (cpf) core_domainsV2AddIdentifier_(state, idPessoa, 'CPF', cpf, false);
  if (email) core_domainsV2AddIdentifier_(state, idPessoa, 'EMAIL', email, false);
  var idProfessor = core_domainsV2GetByAliases_(record, ['ID_PROFESSOR']);
  if (idProfessor) core_domainsV2AddIdentifier_(state, idPessoa, 'ID_PROFESSOR', idProfessor, true);
  var idExterno = core_domainsV2GetByAliases_(record, ['ID_PARTICIPANTE_EXTERNO']);
  if (idExterno) core_domainsV2AddIdentifier_(state, idPessoa, 'ID_PARTICIPANTE_EXTERNO', idExterno, true);

  return idPessoa;
}

function core_domainsV2CreatePessoasState_(dest) {
  var pessoasById = core_domainsV2IndexByField_(dest.PESSOAS_BASE.records, 'ID_PESSOA');
  var personKeyToId = {};
  var identifierKeys = {};
  dest.PESSOAS_IDENTIFICADORES.records.forEach(function(record) {
    var type = String(record.TIPO_IDENTIFICADOR || '').trim();
    var value = String(record.VALOR_IDENTIFICADOR || '').trim();
    var idPessoa = String(record.ID_PESSOA || '').trim();
    var mapKey = core_domainsV2IdentifierMapKey_(type, value);
    if (idPessoa && type && value) identifierKeys[idPessoa + '|' + type + '|' + value] = true;
    if (mapKey && idPessoa) personKeyToId[mapKey] = idPessoa;
  });
  dest.PESSOAS_BASE.records.forEach(function(record) {
    var email = core_domainsV2NormalizeEmail_(record.EMAIL_PRINCIPAL);
    var cpf = core_domainsV2OnlyDigits_(record.CPF);
    if (email && !personKeyToId['EMAIL:' + email]) personKeyToId['EMAIL:' + email] = record.ID_PESSOA;
    if (cpf && !personKeyToId['CPF:' + cpf]) personKeyToId['CPF:' + cpf] = record.ID_PESSOA;
  });

  var memberDetailKeys = {};
  dest.MEMBROS_DETALHES.records.forEach(function(record) {
    var key = String(record.ID_PESSOA || '').trim() || ('RGA:' + String(record.RGA || '').trim());
    if (key) memberDetailKeys[key] = true;
  });
  var colaboradorKeys = {};
  dest.COLABORADORES_ACADEMICOS.records.forEach(function(record) {
    var key = String(record.ID_PESSOA || '').trim() || ('ID_PROFESSOR:' + String(record.ID_PROFESSOR || '').trim());
    if (key) colaboradorKeys[key] = true;
  });
  var externoKeys = {};
  dest.PARTICIPANTES_EXTERNOS_DETALHES.records.forEach(function(record) {
    var key = String(record.ID_PESSOA || '').trim() || ('ID_EXTERNO:' + String(record.ID_PARTICIPANTE_EXTERNO || '').trim());
    if (key) externoKeys[key] = true;
  });
  var eventoKeys = {};
  dest.MEMBROS_EVENTOS_VINCULO.records.forEach(function(record) {
    var key = String(record.ID_EVENTO_MEMBRO || '').trim();
    if (key) eventoKeys[key] = true;
  });
  var resumoKeys = {};
  dest.PESSOAS_RESUMO_OPERACIONAL.records.forEach(function(record) {
    var key = String(record.ID_PESSOA || '').trim();
    if (key) resumoKeys[key] = true;
  });
  var vinculoKeys = {};
  dest.VINCULOS_GEAPA.records.forEach(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    var tipo = String(record.TIPO_VINCULO || '').trim();
    if (idPessoa && tipo) vinculoKeys[idPessoa + '|' + tipo] = true;
  });
  var consentimentoKeys = {};
  dest.PESSOAS_COMUNICACAO_CONSENTIMENTOS.records.forEach(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    var email = core_domainsV2NormalizeEmail_(record.EMAIL);
    var origem = String(record.ORIGEM_CONSENTIMENTO || '').trim();
    if (idPessoa || email || origem) consentimentoKeys[idPessoa + '|' + email + '|' + origem] = true;
  });

  return {
    pessoas: [],
    identificadores: [],
    membrosDetalhes: [],
    colaboradores: [],
    externos: [],
    vinculos: [],
    eventos: [],
    consentimentos: [],
    resumo: [],
    personKeyToId: personKeyToId,
    pessoasById: pessoasById,
    identifierKeys: identifierKeys,
    identifierConflicts: [],
    vinculoKeys: vinculoKeys,
    consentimentoKeys: consentimentoKeys,
    memberDetailKeys: memberDetailKeys,
    colaboradorKeys: colaboradorKeys,
    externoKeys: externoKeys,
    eventoKeys: eventoKeys,
    resumoKeys: resumoKeys,
    nextPessoa: core_domainsV2MaxSeq_(dest.PESSOAS_BASE.records, 'ID_PESSOA', 'PES'),
    nextIdentificador: core_domainsV2MaxSeq_(dest.PESSOAS_IDENTIFICADORES.records, 'ID_IDENTIFICADOR', 'IDE'),
    nextVinculo: core_domainsV2MaxSeq_(dest.VINCULOS_GEAPA.records, 'ID_VINCULO', 'VIN')
  };
}

function core_domainsV2PessoasReport_() {
  return {
    totalMembrosAtuaisLidos: 0,
    totalExMembrosLidos: 0,
    totalMembrosEsperaLidos: 0,
    totalProfessoresTecnicosLidos: 0,
    totalExternosLidos: 0,
    totalPessoasUnicasGeradas: 0,
    totalIdentificadoresGerados: 0,
    totalVinculosGerados: 0,
    totalMembrosDetalhesGerados: 0,
    totalConsentimentosGerados: 0,
    totalEventosMigrados: 0,
    registrosSemRga: 0,
    registrosSemEmail: 0,
    registrosSemNome: 0,
    duplicidadesPorEmail: [],
    duplicidadesPorCpf: [],
    membrosOrdenadosComDataIntegracao: 0,
    membrosOrdenadosSemDataIntegracao: 0,
    conflitosIdentificadores: [],
    fontes: [],
    revisoes: [],
    emailCounts: {},
    cpfCounts: {}
  };
}

function core_domainsV2Limit_(records, limit) {
  return limit == null ? records : records.slice(0, limit);
}

function core_domainsV2ParseIntegrationDate_(record) {
  var value = core_domainsV2GetByAliases_(record, [
    'DATA_INTEGRACAO',
    'DATA_INTEGRACAO_ORIGINAL',
    'DATA_ENTRADA',
    'DATA_INICIO',
    'Data de integração',
    'Data de integracao'
  ]);
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value.getTime();
  }

  var text = String(value || '').trim();
  var iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])).getTime();

  var br = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (br) return new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1])).getTime();

  var parsed = Date.parse(text);
  return isNaN(parsed) ? null : parsed;
}

function core_domainsV2StableMemberSortKey_(record) {
  return [
    core_domainsV2GetRga_(record),
    core_domainsV2NormalizeEmail_(core_domainsV2GetByAliases_(record, ['EMAIL', 'E-mail'])),
    core_domainsV2NormalizeText_(core_domainsV2GetByAliases_(record, ['NOME_MEMBRO', 'NOME_COMPLETO', 'NOME']))
  ].join('|');
}

function core_domainsV2JoinNonEmpty_(values, separator) {
  var seen = {};
  return (values || []).map(function(value) {
    return String(value || '').trim();
  }).filter(function(value) {
    if (!value) return false;
    var key = core_domainsV2NormalizeText_(value);
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  }).join(separator || ' | ');
}

function core_domainsV2CollectEixos_(record) {
  var values = [
    core_domainsV2GetByAliases_(record, ['EIXOS_INTERESSE']),
    core_domainsV2GetByAliases_(record, ['EIXO_ASSOCIADO']),
    core_domainsV2GetByAliases_(record, ['EIXO_ASSOCIADO_1']),
    core_domainsV2GetByAliases_(record, ['EIXO_ASSOCIADO_2']),
    core_domainsV2GetByAliases_(record, ['EIXO_TEMATICO_1', 'EIXO TEMATICO 1']),
    core_domainsV2GetByAliases_(record, ['EIXO_TEMATICO_2', 'EIXO TEMATICO 2'])
  ];

  [
    'INTERESSE_EIXO_I',
    'INTERESSE_EIXO_II',
    'INTERESSE_EIXO_III',
    'INTERESSE_EIXO_IV',
    'INTERESSE_EIXO_V',
    'INTERESSE_EIXO_VI',
    'INTERESSE_EIXO_VII',
    'INTERESSE_EIXO_VIII'
  ].forEach(function(header) {
    var value = core_domainsV2GetByAliases_(record, [header]);
    if (core_domainsV2NormalizeText_(value) === 'SIM') {
      values.push(header.replace('INTERESSE_', ''));
    } else if (value) {
      values.push(value);
    }
  });

  return core_domainsV2JoinNonEmpty_(values, ' | ');
}

function core_domainsV2IsSim_(value) {
  var normalized = core_domainsV2NormalizeText_(value);
  return normalized === 'SIM' || normalized === 'S' || normalized === 'YES' || normalized === 'TRUE';
}

function core_domainsV2IsNao_(value) {
  var normalized = core_domainsV2NormalizeText_(value);
  return normalized === 'NAO' || normalized === 'N' || normalized === 'NO' || normalized === 'FALSE';
}

function core_domainsV2DeriveRecebeComunicacoes_(record) {
  var direct = core_domainsV2GetByAliases_(record, ['RECEBE_COMUNICACOES_GEAPA', 'AUTORIZA_COMUNICACAO']);
  if (direct !== '' && direct != null) return direct;

  var channelValues = [
    core_domainsV2GetByAliases_(record, ['RECEBE_COMUNICADOS_GERAIS']),
    core_domainsV2GetByAliases_(record, ['RECEBE_REUNIOES_ABERTAS']),
    core_domainsV2GetByAliases_(record, ['RECEBE_APRESENTACOES_ALUNOS']),
    core_domainsV2GetByAliases_(record, ['RECEBE_EVENTOS_VISITAS'])
  ].filter(function(value) {
    return value !== '' && value != null;
  });

  if (channelValues.some(core_domainsV2IsSim_)) return 'SIM';
  if (channelValues.length && channelValues.every(core_domainsV2IsNao_)) return 'NAO';
  return '';
}

function core_domainsV2DeriveStatusComunicacao_(record, recebe, email) {
  var direct = core_domainsV2GetByAliases_(record, ['STATUS_COMUNICACAO', 'STATUS_CONTATO']);
  if (direct !== '' && direct != null) return direct;
  if (core_domainsV2IsNao_(recebe)) return 'INATIVO';

  var ativo = core_domainsV2GetByAliases_(record, ['ATIVO', 'STATUS_REGISTRO']);
  if (email && core_domainsV2IsSim_(recebe) && !core_domainsV2IsNao_(ativo)) return 'ATIVO';
  return '';
}

function core_domainsV2EixoTokenMatches_(eixos, roman) {
  roman = String(roman || '').trim().toUpperCase();
  if (!roman) return false;

  return String(eixos || '').split(/[|;,\n]+/).some(function(part) {
    var normalized = core_domainsV2NormalizeText_(part);
    if (!normalized) return false;
    return normalized === roman ||
      normalized === 'EIXO_' + roman ||
      normalized === 'EIXO ' + roman ||
      normalized.indexOf(roman + ' -') === 0 ||
      normalized.indexOf('EIXO ' + roman + ' -') === 0 ||
      normalized.indexOf('EIXO_' + roman + ' -') === 0;
  });
}

function core_domainsV2GetEixoInterestFlag_(record, header, roman, eixos) {
  var explicit = core_domainsV2GetByAliases_(record, [header]);
  if (explicit !== '' && explicit != null) return explicit;
  return core_domainsV2EixoTokenMatches_(eixos, roman) ? 'SIM' : '';
}

function core_domainsV2CommunicationPrefsObs_(record) {
  return core_domainsV2JoinNonEmpty_([
    core_domainsV2GetByAliases_(record, ['RECEBE_COMUNICADOS_GERAIS']) ? 'RECEBE_COMUNICADOS_GERAIS=' + core_domainsV2GetByAliases_(record, ['RECEBE_COMUNICADOS_GERAIS']) : '',
    core_domainsV2GetByAliases_(record, ['RECEBE_REUNIOES_ABERTAS']) ? 'RECEBE_REUNIOES_ABERTAS=' + core_domainsV2GetByAliases_(record, ['RECEBE_REUNIOES_ABERTAS']) : '',
    core_domainsV2GetByAliases_(record, ['RECEBE_APRESENTACOES_ALUNOS']) ? 'RECEBE_APRESENTACOES_ALUNOS=' + core_domainsV2GetByAliases_(record, ['RECEBE_APRESENTACOES_ALUNOS']) : '',
    core_domainsV2GetByAliases_(record, ['RECEBE_EVENTOS_VISITAS']) ? 'RECEBE_EVENTOS_VISITAS=' + core_domainsV2GetByAliases_(record, ['RECEBE_EVENTOS_VISITAS']) : ''
  ], ' | ');
}

function core_domainsV2AddCommunicationConsent_(state, idPessoa, record, sourceLabel) {
  var email = core_domainsV2NormalizeEmail_(core_domainsV2GetByAliases_(record, ['EMAIL', 'EMAIL_PRINCIPAL', 'E-mail']));
  var recebe = core_domainsV2DeriveRecebeComunicacoes_(record);
  var status = core_domainsV2DeriveStatusComunicacao_(record, recebe, email);
  var eixos = core_domainsV2CollectEixos_(record);

  if (!email && !recebe && !status && !eixos) return;

  var key = idPessoa + '|' + email + '|' + sourceLabel;
  if (state.consentimentoKeys[key]) return;
  state.consentimentoKeys[key] = true;
  state.consentimentos.push({
    ID_PESSOA: idPessoa,
    EMAIL: email,
    RECEBE_COMUNICACOES_GEAPA: recebe || '',
    STATUS_COMUNICACAO: status || '',
    EIXOS_INTERESSE: eixos,
    INTERESSE_EIXO_I: core_domainsV2GetEixoInterestFlag_(record, 'INTERESSE_EIXO_I', 'I', eixos),
    INTERESSE_EIXO_II: core_domainsV2GetEixoInterestFlag_(record, 'INTERESSE_EIXO_II', 'II', eixos),
    INTERESSE_EIXO_III: core_domainsV2GetEixoInterestFlag_(record, 'INTERESSE_EIXO_III', 'III', eixos),
    INTERESSE_EIXO_IV: core_domainsV2GetEixoInterestFlag_(record, 'INTERESSE_EIXO_IV', 'IV', eixos),
    INTERESSE_EIXO_V: core_domainsV2GetEixoInterestFlag_(record, 'INTERESSE_EIXO_V', 'V', eixos),
    INTERESSE_EIXO_VI: core_domainsV2GetEixoInterestFlag_(record, 'INTERESSE_EIXO_VI', 'VI', eixos),
    INTERESSE_EIXO_VII: core_domainsV2GetEixoInterestFlag_(record, 'INTERESSE_EIXO_VII', 'VII', eixos),
    INTERESSE_EIXO_VIII: core_domainsV2GetEixoInterestFlag_(record, 'INTERESSE_EIXO_VIII', 'VIII', eixos),
    ORIGEM_CONSENTIMENTO: core_domainsV2GetByAliases_(record, ['ORIGEM_AUTORIZACAO', 'ORIGEM_CONTATO']) || sourceLabel,
    DATA_CONSENTIMENTO: core_domainsV2GetByAliases_(record, ['DATA_AUTORIZACAO_COMUNICACAO', 'DATA_CADASTRO_FORM']),
    DATA_REVOGACAO: core_domainsV2GetByAliases_(record, ['DATA_DESCADASTRAMENTO']),
    OBS: core_domainsV2JoinNonEmpty_([
      core_domainsV2GetByAliases_(record, ['OBS_COMUNICACAO', 'OBSERVACOES', 'OBS']),
      core_domainsV2CommunicationPrefsObs_(record)
    ], ' | ')
  });
}

function core_domainsV2BuildSortedMemberMigrationItems_(sources, limit) {
  var items = [];

  function add(records, sourceLabel, kind, tipoVinculo, statusVinculo, sourceOrder) {
    core_domainsV2Limit_(records, limit).forEach(function(record, index) {
      items.push({
        record: record,
        sourceLabel: sourceLabel,
        kind: kind,
        tipoVinculo: tipoVinculo,
        statusVinculo: statusVinculo,
        integrationTime: core_domainsV2ParseIntegrationDate_(record),
        stableKey: core_domainsV2StableMemberSortKey_(record),
        sourceOrder: sourceOrder,
        rowOrder: index
      });
    });
  }

  add(sources.membrosAtuais, 'Membros Atuais', 'MEMBRO', 'MEMBRO_EFETIVO', 'ATIVO', 1);
  add(sources.exMembros, 'Ex-Membros', 'EGRESSO', 'EGRESSO', 'ATIVO', 2);
  add(sources.membrosEspera, 'Membros em Espera', 'ESPERA', 'MEMBRO_EM_ESPERA', 'EM_ESPERA', 3);

  return items.sort(function(a, b) {
    var aHasDate = a.integrationTime != null;
    var bHasDate = b.integrationTime != null;
    if (aHasDate && bHasDate && a.integrationTime !== b.integrationTime) {
      return a.integrationTime - b.integrationTime;
    }
    if (aHasDate !== bHasDate) return aHasDate ? -1 : 1;
    if (a.stableKey !== b.stableKey) return a.stableKey < b.stableKey ? -1 : 1;
    if (a.sourceOrder !== b.sourceOrder) return a.sourceOrder - b.sourceOrder;
    return a.rowOrder - b.rowOrder;
  });
}

function core_domainsV2ProcessMemberLikeItem_(state, report, item) {
  var record = item.record;
  var sourceLabel = item.sourceLabel;
  var kind = item.kind;
  var tipoVinculo = item.tipoVinculo;
  var statusVinculo = item.statusVinculo;

    var idPessoa = core_domainsV2EnsurePessoa_(state, kind, record, sourceLabel, report);
    if (!idPessoa) return;
    var rga = core_domainsV2GetRga_(record);
    if (!state.memberDetailKeys[idPessoa]) {
      state.memberDetailKeys[idPessoa] = true;
      state.membrosDetalhes.push({
        ID_PESSOA: idPessoa,
        RGA: rga,
        SEMESTRE_ENTRADA: core_domainsV2GetByAliases_(record, ['SEMESTRE_ENTRADA', 'SEMESTRE_INSCRICAO', 'Semestre de Entrada']),
        SEMESTRE_ATUAL: core_domainsV2GetByAliases_(record, ['SEMESTRE_ATUAL', 'Semestre atual']),
        DATA_INTEGRACAO_ORIGINAL: core_domainsV2GetByAliases_(record, ['DATA_INTEGRACAO', 'DATA_ENTRADA']),
        HISTORICO_ATIVIDADES_ACADEMICAS: core_domainsV2GetByAliases_(record, ['HISTORICO_ATIVIDADES_ACADEMICAS']),
        OBS_MEMBRO: core_domainsV2JoinNonEmpty_([
          'Migrado de ' + sourceLabel + '.',
          rga ? '' : 'RGA nao informado/encontrado na fonte legada.'
        ], ' | ')
      });
    }
    core_domainsV2AddVinculo_(state, idPessoa, tipoVinculo, statusVinculo, record, sourceLabel);
    core_domainsV2AddCommunicationConsent_(state, idPessoa, record, sourceLabel);
    if (!state.resumoKeys[idPessoa]) {
      state.resumoKeys[idPessoa] = true;
      state.resumo.push({
      ID_PESSOA: idPessoa,
      RGA: rga,
      NOME_EXIBICAO: core_domainsV2GetByAliases_(record, ['NOME_EXIBICAO', 'NOME_MEMBRO', 'NOME']),
      EMAIL: core_domainsV2NormalizeEmail_(core_domainsV2GetByAliases_(record, ['EMAIL'])),
      TIPO_VINCULO_ATUAL: tipoVinculo,
      STATUS_VINCULO_ATUAL: statusVinculo,
      CARGO_FUNCAO_ATUAL: core_domainsV2GetByAliases_(record, ['CARGO_FUNCAO_ATUAL', 'Cargo/Função atual', 'Ocupação atual']),
      PERFIL_PORTAL_CALCULADO: core_domainsV2GetByAliases_(record, ['PERFIL_PORTAL']),
      PORTAL_ATIVO: core_domainsV2GetByAliases_(record, ['PORTAL_ATIVO']),
      TEMPO_EFETIVO_NO_GRUPO: '',
      QTD_SEMESTRES_NO_GRUPO: core_domainsV2GetByAliases_(record, ['QTD_SEMESTRES_NO_GRUPO']),
      QTD_APRESENTACOES_REALIZADAS: '',
      PERIODO_ULTIMA_APRESENTACAO: '',
      FREQUENCIA_RESUMIDA: '',
      PENDENCIAS_ABERTAS: '',
      FLAG_JA_FOI_SUSPENSO: '',
      STATUS_ELEGIBILIDADE_DIRETORIA: '',
      DATA_LIMITE_ESTIMADA_DIRETORIA: '',
      ULTIMA_ATUALIZACAO: new Date()
      });
    }
}

function coreDiagnosticarMigracaoPessoasV2_() {
  var report = [];
  Object.keys(CORE_DOMAINS_V2_MIGRATION_CFG.pessoasSources).forEach(function(key) {
    var cfg = CORE_DOMAINS_V2_MIGRATION_CFG.pessoasSources[key];
    var found = core_domainsV2FindSheetFromSource_(cfg);
    report.push({
      fonte: cfg.label,
      encontrada: !!found.sheet,
      origem: found.source,
      registryKey: found.key,
      sheetName: found.sheet ? found.sheet.getName() : '',
      registros: found.sheet ? core_domainsV2ReadRecordsSafe_(found.sheet).length : 0,
      erros: found.errors || []
    });
  });
  return {
    ok: report.every(function(item) { return item.encontrada; }),
    fontes: report,
    destino: coreDiagnosticarCentralDomainsV2DevSheets_().diagnostics.filter(function(item) {
      return item.domain === 'PESSOAS';
    })
  };
}

function coreMigrarPessoasV2_(options) {
  var opts = core_domainsV2MigrationOptions_(options);
  coreEnsurePessoasV2DevSheets_({ applyUx: true });
  if (opts.resetDestino && !opts.dryRun) {
    core_domainsV2ClearDataRowsForDomain_('PESSOAS', opts);
    coreEnsurePessoasV2DevSheets_({ applyUx: true });
  }
  var ss = core_domainsV2OpenDestSpreadsheet_('PESSOAS');
  var dest = {};
  CORE_DOMAINS_V2_SCHEMAS.PESSOAS.forEach(function(def) {
    dest[def.sheetName] = core_domainsV2ReadDest_(ss, def.sheetName);
  });

  var sources = {};
  Object.keys(CORE_DOMAINS_V2_MIGRATION_CFG.pessoasSources).forEach(function(key) {
    var found = core_domainsV2FindSheetFromSource_(CORE_DOMAINS_V2_MIGRATION_CFG.pessoasSources[key]);
    sources[key] = found.sheet ? core_domainsV2ReadRecordsSafe_(found.sheet) : [];
  });

  var report = core_domainsV2PessoasReport_();
  report.totalMembrosAtuaisLidos = sources.membrosAtuais.length;
  report.totalExMembrosLidos = sources.exMembros.length;
  report.totalMembrosEsperaLidos = sources.membrosEspera.length;
  report.totalProfessoresTecnicosLidos = sources.professores.length;
  report.totalExternosLidos = sources.externos.length;

  var state = core_domainsV2CreatePessoasState_(dest);
  var memberItems = core_domainsV2BuildSortedMemberMigrationItems_(sources, opts.limit);
  memberItems.forEach(function(item) {
    if (item.integrationTime != null) {
      report.membrosOrdenadosComDataIntegracao++;
    } else {
      report.membrosOrdenadosSemDataIntegracao++;
    }
    core_domainsV2ProcessMemberLikeItem_(state, report, item);
  });

  core_domainsV2Limit_(sources.professores, opts.limit).forEach(function(record) {
    var idPessoa = core_domainsV2EnsurePessoa_(state, 'PROFESSOR', record, 'Dados dos Professores/Tecnicos', report);
    if (!idPessoa) return;
    if (state.colaboradorKeys[idPessoa]) return;
    state.colaboradorKeys[idPessoa] = true;
    var eixoAssociado = core_domainsV2JoinNonEmpty_([
      core_domainsV2GetByAliases_(record, ['EIXO_ASSOCIADO']),
      core_domainsV2GetByAliases_(record, ['EIXO_ASSOCIADO_1', 'EIXO_TEMATICO_1', 'EIXO TEMATICO 1']),
      core_domainsV2GetByAliases_(record, ['EIXO_ASSOCIADO_2', 'EIXO_TEMATICO_2', 'EIXO TEMATICO 2'])
    ], ' | ');
    state.colaboradores.push({
      ID_PESSOA: idPessoa,
      ID_PROFESSOR: core_domainsV2GetByAliases_(record, ['ID_PROFESSOR']),
      TIPO_COLABORADOR: core_domainsV2GetByAliases_(record, ['TIPO_COLABORADOR', 'TIPO_PROFISSIONAL', 'TIPO', 'CATEGORIA']),
      INSTITUICAO: core_domainsV2GetByAliases_(record, ['INSTITUICAO', 'INSTITUIÇÃO']),
      SETOR: core_domainsV2GetByAliases_(record, ['SETOR', 'CURSO_VINCULO']),
      TITULACAO: core_domainsV2GetByAliases_(record, ['TITULACAO', 'TITULAÇÃO']),
      AREA_ATUACAO: core_domainsV2GetByAliases_(record, ['AREA_ATUACAO', 'ÁREA_ATUAÇÃO', 'DISCIPLINAS_AREAS']),
      EIXO_ASSOCIADO: eixoAssociado,
      EMAIL_INSTITUCIONAL: core_domainsV2NormalizeEmail_(core_domainsV2GetByAliases_(record, ['EMAIL_INSTITUCIONAL', 'EMAIL', 'E-mail'])),
      LINK_LATTES: core_domainsV2GetByAliases_(record, ['LINK_LATTES', 'LATTES', 'CURRICULO_LATTES']),
      OBS_ACADEMICA: core_domainsV2JoinNonEmpty_([
        'Migrado de Dados dos Professores/Tecnicos.',
        core_domainsV2GetByAliases_(record, ['VINCULO_GEAPA']) ? 'VINCULO_GEAPA=' + core_domainsV2GetByAliases_(record, ['VINCULO_GEAPA']) : '',
        core_domainsV2GetByAliases_(record, ['FORMACAO']) ? 'FORMACAO=' + core_domainsV2GetByAliases_(record, ['FORMACAO']) : '',
        core_domainsV2GetByAliases_(record, ['OBSERVACOES', 'OBS'])
      ], ' | '),
      ATIVO: 'SIM'
    });
    var vinculoGeapa = core_domainsV2NormalizeText_(core_domainsV2GetByAliases_(record, ['VINCULO_GEAPA']));
    if (vinculoGeapa && vinculoGeapa !== 'SEM_VINCULO_ATIVO') {
      core_domainsV2AddVinculo_(state, idPessoa, vinculoGeapa, 'ATIVO', record, 'Dados dos Professores/Tecnicos');
    }
    core_domainsV2AddCommunicationConsent_(state, idPessoa, record, 'Dados dos Professores/Tecnicos');
  });

  core_domainsV2Limit_(sources.externos, opts.limit).forEach(function(record) {
    var idPessoa = core_domainsV2EnsurePessoa_(state, 'EXTERNO', record, 'Participantes Externos', report);
    if (!idPessoa) return;
    if (state.externoKeys[idPessoa]) return;
    state.externoKeys[idPessoa] = true;
    state.externos.push({
      ID_PESSOA: idPessoa,
      ID_PARTICIPANTE_EXTERNO: core_domainsV2GetByAliases_(record, ['ID_PARTICIPANTE_EXTERNO']),
      TIPO_PUBLICO: core_domainsV2GetByAliases_(record, ['TIPO_PUBLICO', 'CATEGORIA_PUBLICO', 'TIPO']),
      INSTITUICAO_ORIGEM: core_domainsV2GetByAliases_(record, ['INSTITUICAO_ORIGEM', 'INSTITUICAO']),
      EMPRESA_ORIGEM: core_domainsV2GetByAliases_(record, ['EMPRESA_ORIGEM']),
      CARGO_PROFISSAO: core_domainsV2GetByAliases_(record, ['CARGO_PROFISSAO', 'CARGO_OU_ATUACAO']),
      CURSO: core_domainsV2GetByAliases_(record, ['CURSO', 'CURSO_OU_AREA']),
      CIDADE_ATUAL: core_domainsV2GetByAliases_(record, ['CIDADE_ATUAL', 'CIDADE']),
      EVENTO_ORIGEM: core_domainsV2GetByAliases_(record, ['EVENTO_ORIGEM', 'ORIGEM_CONTATO']),
      AUTORIZA_CONTATO: core_domainsV2GetByAliases_(record, ['AUTORIZA_CONTATO', 'AUTORIZA_COMUNICACAO', 'AUTORIZA_ARMAZENAMENTO_DADOS']),
      OBS_EXTERNA: core_domainsV2JoinNonEmpty_([
        'Migrado de Participantes Externos.',
        core_domainsV2GetByAliases_(record, ['RELACAO_COM_GEAPA']) ? 'RELACAO_COM_GEAPA=' + core_domainsV2GetByAliases_(record, ['RELACAO_COM_GEAPA']) : '',
        core_domainsV2GetByAliases_(record, ['MOTIVACAO_OU_INTERESSE']) ? 'MOTIVACAO_OU_INTERESSE=' + core_domainsV2GetByAliases_(record, ['MOTIVACAO_OU_INTERESSE']) : '',
        core_domainsV2GetByAliases_(record, ['OBSERVACOES', 'OBS'])
      ], ' | '),
      ATIVO: 'SIM'
    });
    core_domainsV2AddCommunicationConsent_(state, idPessoa, record, 'Participantes Externos');
  });

  core_domainsV2Limit_(sources.eventos, opts.limit).forEach(function(record) {
    var eventId = core_domainsV2GetByAliases_(record, ['ID_EVENTO_MEMBRO', 'ID_EVENTO']) || ('EVT-LEG-' + record.__rowNumber);
    if (state.eventoKeys[eventId]) return;
    state.eventoKeys[eventId] = true;
    state.eventos.push({
      ID_EVENTO_MEMBRO: eventId,
      RGA: core_domainsV2GetByAliases_(record, ['RGA']),
      ID_PESSOA: '',
      TIPO_EVENTO: core_domainsV2GetByAliases_(record, ['TIPO_EVENTO', 'TIPO']),
      DATA_EVENTO: core_domainsV2GetByAliases_(record, ['DATA_EVENTO', 'DATA']),
      STATUS_EVENTO: core_domainsV2GetByAliases_(record, ['STATUS_EVENTO', 'STATUS']),
      MODULO_ORIGEM: core_domainsV2GetByAliases_(record, ['MODULO_ORIGEM']),
      CHAVE_ORIGEM: core_domainsV2GetByAliases_(record, ['CHAVE_ORIGEM']),
      OBSERVACOES: core_domainsV2GetByAliases_(record, ['OBSERVACOES', 'OBS']),
      ATUALIZADO_EM: new Date(),
      PROCESSADO_POR_MODULO: core_domainsV2GetByAliases_(record, ['PROCESSADO_POR_MODULO']),
      DATA_PROCESSAMENTO: core_domainsV2GetByAliases_(record, ['DATA_PROCESSAMENTO']),
      ERRO_PROCESSAMENTO: core_domainsV2GetByAliases_(record, ['ERRO_PROCESSAMENTO'])
    });
  });

  report.duplicidadesPorEmail = Object.keys(report.emailCounts).filter(function(key) { return report.emailCounts[key] > 1; });
  report.duplicidadesPorCpf = Object.keys(report.cpfCounts).filter(function(key) { return report.cpfCounts[key] > 1; });
  report.totalPessoasUnicasGeradas = state.pessoas.length;
  report.totalIdentificadoresGerados = state.identificadores.length;
  report.totalVinculosGerados = state.vinculos.length;
  report.totalMembrosDetalhesGerados = state.membrosDetalhes.length;
  report.totalConsentimentosGerados = state.consentimentos.length;
  report.totalEventosMigrados = state.eventos.length;
  report.conflitosIdentificadores = state.identifierConflicts;

  if (opts.resetDestino && memberItems.length > 0 && state.membrosDetalhes.length === 0) {
    report.revisoes.push({
      tipo: 'MEMBROS_DETALHES_ZERO_APOS_RESET_E_MIGRACAO',
      mensagem: 'Foram lidos registros de membros, mas nenhuma linha de MEMBROS_DETALHES foi gerada.'
    });
    throw new Error('Migracao Pessoas v2 interrompida: MEMBROS_DETALHES gerou zero linhas apesar de haver membros lidos. Verifique cabecalhos/chaves das fontes legadas.');
  }

  core_domainsV2AppendRows_(dest.PESSOAS_BASE.sheet, state.pessoas, opts.dryRun);
  core_domainsV2AppendRows_(dest.PESSOAS_IDENTIFICADORES.sheet, state.identificadores, opts.dryRun);
  core_domainsV2AppendRows_(dest.MEMBROS_DETALHES.sheet, state.membrosDetalhes, opts.dryRun);
  core_domainsV2AppendRows_(dest.COLABORADORES_ACADEMICOS.sheet, state.colaboradores, opts.dryRun);
  core_domainsV2AppendRows_(dest.PARTICIPANTES_EXTERNOS_DETALHES.sheet, state.externos, opts.dryRun);
  core_domainsV2AppendRows_(dest.VINCULOS_GEAPA.sheet, state.vinculos, opts.dryRun);
  core_domainsV2AppendRows_(dest.PESSOAS_COMUNICACAO_CONSENTIMENTOS.sheet, state.consentimentos, opts.dryRun);
  core_domainsV2AppendRows_(dest.MEMBROS_EVENTOS_VINCULO.sheet, state.eventos, opts.dryRun);
  core_domainsV2AppendRows_(dest.PESSOAS_RESUMO_OPERACIONAL.sheet, state.resumo, opts.dryRun);

  return {
    ok: true,
    dryRun: opts.dryRun,
    destino: CORE_DOMAINS_V2_DEV_SPREADSHEETS.PESSOAS,
    report: report
  };
}

function coreDiagnosticarMigracaoVigenciasV2_() {
  var report = [];
  Object.keys(CORE_DOMAINS_V2_MIGRATION_CFG.vigenciasSources).forEach(function(key) {
    var cfg = CORE_DOMAINS_V2_MIGRATION_CFG.vigenciasSources[key];
    var found = core_domainsV2FindSheetFromSource_(cfg);
    report.push({
      fonte: cfg.label,
      encontrada: !!found.sheet,
      origem: found.source,
      registryKey: found.key,
      sheetName: found.sheet ? found.sheet.getName() : '',
      registros: found.sheet ? core_domainsV2ReadRecordsSafe_(found.sheet).length : 0,
      erros: found.errors || []
    });
  });
  return {
    ok: report.every(function(item) { return item.encontrada; }),
    fontes: report,
    destino: coreDiagnosticarCentralDomainsV2DevSheets_().diagnostics.filter(function(item) {
      return item.domain === 'VIGENCIAS';
    })
  };
}

function core_domainsV2SimpleCopyRows_(records, targetHeaders, mapping, limit) {
  return core_domainsV2Limit_(records, limit).map(function(record) {
    var out = {};
    targetHeaders.forEach(function(header) {
      var aliases = mapping[header] || [header];
      out[header] = core_domainsV2GetByAliases_(record, aliases);
    });
    return out;
  });
}

function core_domainsV2FilterNewRows_(rows, existingRecords, primaryField) {
  var existing = {};
  existingRecords.forEach(function(record) {
    var key = String(record[primaryField] || '').trim();
    if (key) existing[key] = true;
  });
  return rows.filter(function(row) {
    var key = String(row[primaryField] || '').trim();
    if (!key) return true;
    if (existing[key]) return false;
    existing[key] = true;
    return true;
  });
}

function coreMigrarVigenciasV2_(options) {
  var opts = core_domainsV2MigrationOptions_(options);
  coreEnsureVigenciasV2DevSheets_({ applyUx: true });
  if (opts.resetDestino && !opts.dryRun) {
    core_domainsV2ClearDataRowsForDomain_('VIGENCIAS', opts);
  }
  var ss = core_domainsV2OpenDestSpreadsheet_('VIGENCIAS');
  var dest = {};
  CORE_DOMAINS_V2_SCHEMAS.VIGENCIAS.forEach(function(def) {
    dest[def.sheetName] = core_domainsV2ReadDest_(ss, def.sheetName);
  });

  var sourceRecords = {};
  Object.keys(CORE_DOMAINS_V2_MIGRATION_CFG.vigenciasSources).forEach(function(key) {
    var found = core_domainsV2FindSheetFromSource_(CORE_DOMAINS_V2_MIGRATION_CFG.vigenciasSources[key]);
    sourceRecords[key] = found.sheet ? core_domainsV2ReadRecordsSafe_(found.sheet) : [];
  });

  var report = {
    totalSemestresLidos: sourceRecords.semestres.length,
    totalPeriodosLidos: sourceRecords.periodos.length,
    totalDiretoriasLidas: sourceRecords.diretorias.length,
    totalCargosConfigLidos: sourceRecords.cargosConfig.length,
    totalVigenciasGeradas: 0,
    cargosSemCargoKeyCorrespondente: [],
    diretoriasNaoInferidas: [],
    janelasDiretoriaSemIdDiretoria: [],
    revisoes: []
  };

  var semestres = core_domainsV2FilterNewRows_(
    core_domainsV2SimpleCopyRows_(sourceRecords.semestres, CORE_DOMAINS_V2_SCHEMAS.VIGENCIAS[0].headers, {
      ID_SEMESTRE: ['ID_SEMESTRE', 'ID_Semestre'],
      ID_PERIODO: ['ID_PERIODO', 'ID_PerÃ­odo', 'ID_Periodo'],
      NUMERO_REUNIOES_PREVISTAS: ['NUMERO_REUNIOES_PREVISTAS', 'NÃºmero reuniÃµes previstas', 'Numero reunioes previstas'],
      INICIO_MATRICULAS_ONLINE: ['INICIO_MATRICULAS_ONLINE', 'InÃ­cio MatrÃ­culas Online', 'Inicio Matriculas Online'],
      FIM_MATRICULAS_ONLINE: ['FIM_MATRICULAS_ONLINE', 'Fim MatrÃ­culas Online', 'Fim Matriculas Online'],
      INICIO_AJUSTE_ALUNO: ['INICIO_AJUSTE_ALUNO', 'InÃ­cio Ajuste do Aluno', 'Inicio Ajuste do Aluno'],
      FIM_AJUSTE_ALUNO: ['FIM_AJUSTE_ALUNO', 'Fim Ajuste do Aluno'],
      INICIO_AJUSTE_COORDENADOR: ['INICIO_AJUSTE_COORDENADOR', 'InÃ­cio Ajuste do Coordenador', 'Inicio Ajuste do Coordenador'],
      FIM_AJUSTE_COORDENADOR: ['FIM_AJUSTE_COORDENADOR', 'Fim ajuste do Coordenador', 'Fim Ajuste do Coordenador'],
      DATA_INICIO: ['DATA_INICIO', 'Início', 'Inicio'],
      DATA_FIM: ['DATA_FIM', 'Fim'],
      STATUS: ['STATUS'],
      OBS: ['OBS', 'Observacao', 'Observação']
    }, opts.limit),
    dest.SEMESTRES.records,
    'ID_SEMESTRE'
  );
  var periodos = core_domainsV2FilterNewRows_(
    core_domainsV2SimpleCopyRows_(sourceRecords.periodos, CORE_DOMAINS_V2_SCHEMAS.VIGENCIAS[1].headers, {
      ID_PERIODO: ['ID_PERIODO', 'ID_Período', 'ID_Periodo'],
      NOME_PERIODO: ['NOME_PERIODO', 'ID_Período', 'ID_Periodo'],
      TIPO_PERIODO: ['TIPO_PERIODO'],
      NUMERO_MEMBROS_PREVISTOS: ['NUMERO_MEMBROS_PREVISTOS', 'NÃºmero de membros previstos', 'Numero de membros previstos'],
      TOTAL_ATIVIDADES_QUE_CONTAM_FALTA_PLANEJADAS: ['TOTAL_ATIVIDADES_QUE_CONTAM_FALTA_PLANEJADAS'],
      LIMITE_FALTAS_PERIODO_CONGELADO: ['LIMITE_FALTAS_PERIODO_CONGELADO'],
      DATA_FECHAMENTO_PLANEJAMENTO: ['DATA_FECHAMENTO_PLANEJAMENTO'],
      NUMERO_SEI: ['NUMERO_SEI'],
      ANO_SEI: ['ANO_SEI'],
      COORDENADOR: ['COORDENADOR'],
      DATA_INICIO: ['DATA_INICIO', 'Início', 'Inicio'],
      DATA_FIM: ['DATA_FIM', 'Fim'],
      STATUS: ['STATUS'],
      OBS: ['OBS', 'NUMERO_SEI', 'ANO_SEI', 'COORDENADOR']
    }, opts.limit),
    dest.PERIODOS.records,
    'ID_PERIODO'
  );
  var diretorias = core_domainsV2FilterNewRows_(core_domainsV2SimpleCopyRows_(sourceRecords.diretorias, CORE_DOMAINS_V2_SCHEMAS.VIGENCIAS[2].headers, {
    ID_DIRETORIA: ['ID_DIRETORIA', 'ID_Diretoria'],
    NOME_GESTAO: ['NOME_GESTAO', 'NOME', 'Ordem_Diretoria'],
    DATA_INICIO: ['DATA_INICIO', 'Início_Mandato', 'Inicio_Mandato'],
    DATA_FIM_PREVISTA: ['DATA_FIM_PREVISTA', 'Fim_Mandato'],
    DATA_FIM_REAL: ['DATA_FIM_REAL'],
    STATUS_DIRETORIA: ['STATUS_DIRETORIA', 'STATUS'],
    LEMA: ['LEMA', 'SLOGAN', 'Slogan'],
    OBS: ['OBS', 'Plano de trabalho', 'Metas']
  }, opts.limit), dest.DIRETORIAS.records, 'ID_DIRETORIA');
  var semestresDiretoria = core_domainsV2FilterNewRows_(core_domainsV2SimpleCopyRows_(sourceRecords.semestresDiretoria, CORE_DOMAINS_V2_SCHEMAS.VIGENCIAS[3].headers, {
    ID_SEMESTRE_DIRETORIA: ['ID_SEMESTRE_DIRETORIA', 'ID_Janela'],
    ID_DIRETORIA: ['ID_DIRETORIA', 'ID_Diretoria'],
    DATA_INICIO: ['DATA_INICIO', 'Data_Inicio'],
    DATA_FIM: ['DATA_FIM', 'Data_Fim'],
    TOTAL_DIAS: ['TOTAL_DIAS', 'Total_Dias'],
    OBS: ['OBS', 'Observacao', 'Total_Dias']
  }, opts.limit), dest.SEMESTRES_DIRETORIA.records, 'ID_SEMESTRE_DIRETORIA');
  semestresDiretoria.forEach(function(row) {
    if (!row.ID_DIRETORIA) report.janelasDiretoriaSemIdDiretoria.push(row.ID_SEMESTRE_DIRETORIA || '(sem id)');
  });
  var cargosConfig = core_domainsV2FilterNewRows_(core_domainsV2SimpleCopyRows_(sourceRecords.cargosConfig, CORE_DOMAINS_V2_SCHEMAS.VIGENCIAS[4].headers, {
    CARGO_KEY: ['CARGO_KEY'],
    CARGO_NOME: ['CARGO_NOME', 'NOME_PUBLICO'],
    NOME_PUBLICO: ['NOME_PUBLICO', 'CARGO_NOME'],
    TIPO_FUNCAO: ['TIPO_FUNCAO', 'BASE_EXISTENCIA'],
    GRUPO_FUNCAO: ['GRUPO_FUNCAO', 'GRUPO_CARGO'],
    GRUPO_CARGO: ['GRUPO_CARGO', 'GRUPO_FUNCAO'],
    ESCRITA_VARIACAO: ['ESCRITA_VARIACAO'],
    EMAILS_GRUPO: ['EMAILS_GRUPO'],
    OBRIGATORIO_COMPOSICAO_INICIAL: ['OBRIGATORIO_COMPOSICAO_INICIAL'],
    RECEBE_EMAILS: ['RECEBE_EMAILS'],
    E_CARGO_UNICO: ['E_CARGO_UNICO', 'Ã‰_CARGO_UNICO'],
    PERMITIR_NOMEACAO_VIA_FORM: ['PERMITIR_NOMEACAO_VIA_FORM'],
    CONTA_LIMITE_DIRETORIA: ['CONTA_LIMITE_DIRETORIA', 'CONTA_PARA_LIMITE_DIRETORIA'],
    CONTA_PARA_LIMITE_DIRETORIA: ['CONTA_PARA_LIMITE_DIRETORIA', 'CONTA_LIMITE_DIRETORIA'],
    DIREITO_A_VOTO_DIRETORIA: ['DIREITO_A_VOTO_DIRETORIA'],
    EXIGIR_HOMOLOGACAO_PREVIA: ['EXIGIR_HOMOLOGACAO_PREVIA'],
    NIVEL_HIERARQUICO: ['NIVEL_HIERARQUICO', 'DISPLAY_ORDEM'],
    BASE_NORMATIVA: ['BASE_NORMATIVA', 'BASE_EXISTENCIA'],
    ATIVO: ['ATIVO'],
    OBS: ['OBS', 'DIREITO_A_VOTO_DIRETORIA', 'EXIGIR_HOMOLOGACAO_PREVIA', 'OBRIGATORIO_COMPOSICAO_INICIAL']
  }, opts.limit), dest.CARGOS_CONFIG.records, 'CARGO_KEY');
  cargosConfig.forEach(function(row) {
    if (!row.CARGO_KEY) report.cargosSemCargoKeyCorrespondente.push(row.CARGO_NOME || '(sem nome)');
  });

  function buildFuncoes(records, tipo) {
    return core_domainsV2Limit_(records, opts.limit).map(function(record, idx) {
      var cargoKey = core_domainsV2GetByAliases_(record, ['CARGO_KEY', 'Cargo_Key', 'CARGO', 'FUNCAO', 'Cargo/Função', 'Cargo/Funcao']);
      if (!cargoKey) report.cargosSemCargoKeyCorrespondente.push(tipo + ': linha ' + (record.__rowNumber || idx + 2));
      var cargoNomeSnapshot = core_domainsV2GetByAliases_(record, [
        'CARGO_NOME',
        'FUNCAO',
        'Cargo/Função',
        'Cargo/Funcao',
        'Cargo/Função atual',
        'Cargo/Funcao atual',
        'CARGO'
      ]);
      return {
        ID_VIGENCIA: 'VIG-LEG-' + tipo + '-' + (record.__rowNumber || idx + 2),
        ID_PESSOA: '',
        ID_VINCULO: '',
        TIPO_FUNCAO: tipo,
        CARGO_KEY: cargoKey,
        CARGO_NOME_SNAPSHOT: cargoNomeSnapshot || cargoKey,
        ID_DIRETORIA: core_domainsV2GetByAliases_(record, ['ID_DIRETORIA', 'ID_Diretoria']),
        ID_SEMESTRE_DIRETORIA: core_domainsV2GetByAliases_(record, ['ID_SEMESTRE_DIRETORIA', 'ID_Janela']),
        DATA_INICIO: core_domainsV2GetByAliases_(record, ['DATA_INICIO', 'Data_Inicio']),
        DATA_FIM_PREVISTA: core_domainsV2GetByAliases_(record, ['DATA_FIM_PREVISTA', 'Data_Fim_previsto']),
        DATA_FIM_REAL: core_domainsV2GetByAliases_(record, ['DATA_FIM_REAL', 'Data_Fim']),
        STATUS_VIGENCIA: core_domainsV2GetByAliases_(record, ['STATUS_VIGENCIA', 'STATUS']) || 'PENDENTE_REVISAO',
        FONTE_NOMEACAO: 'MIGRACAO_V2_' + tipo,
        LINK_ATA: core_domainsV2GetByAliases_(record, ['LINK_ATA']),
        CRIADO_POR: 'MIGRACAO_V2',
        CRIADO_EM: new Date(),
        ATUALIZADO_POR: 'MIGRACAO_V2',
        ATUALIZADO_EM: new Date(),
        OBS: 'ID_PESSOA deve ser reconciliado pela migracao de Pessoas usando RGA/e-mail.',
        ATIVO: 'SIM'
      };
    });
  }

  var funcoes = core_domainsV2FilterNewRows_([]
    .concat(buildFuncoes(sourceRecords.diretores, 'DIRETORIA'))
    .concat(buildFuncoes(sourceRecords.assessores, 'ASSESSORIA'))
    .concat(buildFuncoes(sourceRecords.conselheiros, 'CONSELHO')), dest.VIGENCIAS_FUNCOES.records, 'ID_VIGENCIA');
  report.totalVigenciasGeradas = funcoes.length;

  core_domainsV2AppendRows_(dest.SEMESTRES.sheet, semestres, opts.dryRun);
  core_domainsV2AppendRows_(dest.PERIODOS.sheet, periodos, opts.dryRun);
  core_domainsV2AppendRows_(dest.DIRETORIAS.sheet, diretorias, opts.dryRun);
  core_domainsV2AppendRows_(dest.SEMESTRES_DIRETORIA.sheet, semestresDiretoria, opts.dryRun);
  core_domainsV2AppendRows_(dest.CARGOS_CONFIG.sheet, cargosConfig, opts.dryRun);
  core_domainsV2AppendRows_(dest.VIGENCIAS_FUNCOES.sheet, funcoes, opts.dryRun);

  return {
    ok: true,
    dryRun: opts.dryRun,
    destino: CORE_DOMAINS_V2_DEV_SPREADSHEETS.VIGENCIAS,
    report: report
  };
}

function coreMigrarDominiosCentraisV2_(options) {
  var opts = core_domainsV2MigrationOptions_(options);
  return {
    dryRun: opts.dryRun,
    pessoas: coreMigrarPessoasV2_(opts),
    vigencias: coreMigrarVigenciasV2_(opts)
  };
}

function coreResetPessoasV2DevDestino_(options) {
  return core_domainsV2ClearDataRowsForDomain_('PESSOAS', options || {});
}

function coreResetVigenciasV2DevDestino_(options) {
  return core_domainsV2ClearDataRowsForDomain_('VIGENCIAS', options || {});
}

function coreResetAndMigrarVigenciasV2_(options) {
  options = options || {};
  var opts = {
    dryRun: false,
    resetDestino: true,
    confirmacao: String(options.confirmacao || '').trim(),
    escreverRelatorio: options.escreverRelatorio !== false
  };
  if (options.limit != null && options.limit !== '') {
    opts.limit = options.limit;
  }
  return coreMigrarVigenciasV2_(opts);
}
