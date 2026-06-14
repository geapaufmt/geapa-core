/**
 * ============================================================
 * 33_core_v2_pessoas_vigencias_rotinas.js
 * ============================================================
 *
 * Rotinas manuais e seguras de Pessoas v2 e Vigencias v2.
 *
 * Regras desta etapa:
 * - usa Registry como fonte unica de resolucao de abas;
 * - nao cria triggers;
 * - dryRun e o padrao para rotinas de atualizacao;
 * - nao apaga, renomeia ou reordena abas;
 * - atualiza caches por cabecalho e preserva colunas extras.
 */

var CORE_V2_ROTINAS_KEYS = Object.freeze({
  PESSOAS: Object.freeze({
    BASE: 'PESSOAS_V2_BASE',
    IDENTIFICADORES: 'PESSOAS_V2_IDENTIFICADORES',
    MEMBROS_DETALHES: 'PESSOAS_V2_MEMBROS_DETALHES',
    VINCULOS: 'PESSOAS_V2_VINCULOS_GEAPA',
    EVENTOS: 'PESSOAS_V2_MEMBROS_EVENTOS_VINCULO',
    RESUMO: 'PESSOAS_V2_RESUMO_OPERACIONAL'
  }),
  VIGENCIAS: Object.freeze({
    SEMESTRES: 'VIGENCIAS_V2_SEMESTRES',
    PERIODOS: 'VIGENCIAS_V2_PERIODOS',
    DIRETORIAS: 'VIGENCIAS_V2_DIRETORIAS',
    SEMESTRES_DIRETORIA: 'VIGENCIAS_V2_SEMESTRES_DIRETORIA',
    CARGOS: 'VIGENCIAS_V2_CARGOS_CONFIG',
    FUNCOES: 'VIGENCIAS_V2_FUNCOES',
    RESUMO: 'VIGENCIAS_V2_RESUMO_ATUAL'
  })
});

var CORE_V2_ROTINAS_FLUXOS = Object.freeze({
  PESSOAS_ATUALIZACAO: Object.freeze({ modulo: 'PESSOAS', fluxo: 'ATUALIZACAO_V2' }),
  PESSOAS_CONFERENCIA: Object.freeze({ modulo: 'PESSOAS', fluxo: 'CONFERENCIA_V2' }),
  VIGENCIAS_ATUALIZACAO: Object.freeze({ modulo: 'VIGENCIAS', fluxo: 'ATUALIZACAO_V2' }),
  VIGENCIAS_CONFERENCIA: Object.freeze({ modulo: 'VIGENCIAS', fluxo: 'CONFERENCIA_V2' })
});

var CORE_V2_ROTINAS_PESSOAS_RESUMO_HEADERS = Object.freeze([
  'ID_PESSOA',
  'NOME_EXIBICAO',
  'EMAIL_PRINCIPAL',
  'EMAIL',
  'RGA',
  'CPF',
  'TIPO_VINCULO_ATUAL',
  'STATUS_VINCULO_ATUAL',
  'DATA_INICIO_VINCULO',
  'DATA_FIM_VINCULO',
  'PORTAL_ATIVO',
  'PERFIL_PORTAL_BASE',
  'PERFIL_PORTAL_CALCULADO',
  'CARGO_FUNCAO_ATUAL',
  'ULTIMA_ATUALIZACAO',
  'OBS_RESUMO'
]);

var CORE_V2_ROTINAS_VIGENCIAS_RESUMO_HEADERS = Object.freeze([
  'ID_VIGENCIA',
  'ID_PESSOA',
  'NOME_EXIBICAO',
  'RGA',
  'OCUPACAO',
  'GRUPO_FUNCAO',
  'DATA_INICIO',
  'DATA_FIM_PREVISTA',
  'DATA_FIM_REAL',
  'STATUS_VIGENCIA',
  'PERFIL_PORTAL_GERADO',
  'PERMISSOES_GERADAS',
  'CARGO_ATUAL_VISIVEL',
  'APARECE_DIRETORIA_PUBLICA',
  'CARGO_FUNCAO_ATUAL',
  'TIPO_FUNCAO_ATUAL',
  'GRUPO_FUNCAO_ATUAL',
  'ID_DIRETORIA_ATUAL',
  'PERFIS_PORTAL_CALCULADOS',
  'PERMISSOES_CALCULADAS',
  'DATA_INICIO_FUNCAO_ATUAL',
  'ULTIMA_ATUALIZACAO'
]);

function core_v2RotinasOptions_(options) {
  options = options || {};
  return {
    dryRun: options.dryRun !== false,
    includeRecords: options.includeRecords === true,
    limit: Math.max(1, Number(options.limit || 10)),
    ambiente: options.ambiente || '',
    allowWriteWithErrors: options.allowWriteWithErrors === true
  };
}

function core_v2RotinasNormalize_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  }).replace(/\s+/g, '_');
}

function core_v2RotinasText_(value) {
  return String(value == null ? '' : value).trim();
}

function core_v2RotinasEmail_(value) {
  return core_normalizeEmail_(value);
}

function core_v2RotinasDigits_(value) {
  return core_onlyDigits_(value);
}

function core_v2RotinasIsSim_(value) {
  var normalized = core_v2RotinasNormalize_(value);
  return normalized === 'SIM' || normalized === 'S' || normalized === 'TRUE' || normalized === 'ATIVO' || normalized === 'ATIVA';
}

function core_v2RotinasIsNao_(value) {
  var normalized = core_v2RotinasNormalize_(value);
  return normalized === 'NAO' || normalized === 'N' || normalized === 'FALSE' || normalized === 'INATIVO' || normalized === 'INATIVA';
}

function core_v2RotinasStatus_(value) {
  return core_v2RotinasNormalize_(value);
}

function core_v2RotinasTipoVinculo_(value) {
  var tipo = core_v2RotinasStatus_(value);
  if (tipo === 'EX_MEMBRO' || tipo === 'EX-MEMBRO' || tipo === 'EX_MEMBROS' || tipo === 'EX MEMBRO') return 'EGRESSO';
  return tipo;
}

function core_v2RotinasIsRecordActive_(record, statusField, activeField) {
  record = record || {};
  var status = core_v2RotinasStatus_(record[statusField || 'STATUS']);
  if (status === 'ATIVO' || status === 'ATIVA' || status === 'VIGENTE') return true;
  if (status === 'INATIVO' || status === 'INATIVA' || status === 'ENCERRADO' || status === 'ENCERRADA' || status === 'CANCELADO' || status === 'CANCELADA') return false;
  if (Object.prototype.hasOwnProperty.call(record, activeField || 'ATIVO')) {
    return core_v2RotinasIsSim_(record[activeField || 'ATIVO']);
  }
  return false;
}

function core_v2RotinasDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  var text = core_v2RotinasText_(value);
  if (!text) return null;
  var m = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    var year = Number(m[3]);
    if (year < 100) year += 2000;
    return new Date(year, Number(m[2]) - 1, Number(m[1]));
  }
  var parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function core_v2RotinasIntervalsOverlap_(aStart, aEnd, bStart, bEnd) {
  if (!aStart || !bStart) return false;
  var aLast = aEnd || new Date(8640000000000000);
  var bLast = bEnd || new Date(8640000000000000);
  return aStart <= bLast && bStart <= aLast;
}

function core_v2RotinasClone_(record) {
  var out = {};
  Object.keys(record || {}).forEach(function(key) {
    if (key !== '__rowNumber') out[key] = record[key];
  });
  return out;
}

function core_v2RotinasIndexFirstBy_(records, field) {
  var out = {};
  (records || []).forEach(function(record) {
    var key = core_v2RotinasText_(record[field]);
    if (key && !out[key]) out[key] = record;
  });
  return out;
}

function core_v2RotinasIndexManyBy_(records, field) {
  var out = {};
  (records || []).forEach(function(record) {
    var key = core_v2RotinasText_(record[field]);
    if (!key) return;
    if (!out[key]) out[key] = [];
    out[key].push(record);
  });
  return out;
}

function core_v2RotinasNewConsistencyEnvelope_(modulo, fluxo) {
  return {
    ok: true,
    modulo: modulo,
    fluxo: fluxo,
    totalVerificado: 0,
    totalInconsistencias: 0,
    inconsistencias: []
  };
}

function core_v2RotinasAddInconsistency_(envelope, gravidade, entidade, idEntidade, campo, valorAtual, regra, mensagem, acaoRecomendada) {
  envelope.inconsistencias.push({
    gravidade: gravidade,
    entidade: entidade,
    idEntidade: idEntidade || '',
    campo: campo || '',
    valorAtual: valorAtual == null ? '' : valorAtual,
    regra: regra || '',
    mensagem: mensagem || '',
    acaoRecomendada: acaoRecomendada || ''
  });
  envelope.totalInconsistencias = envelope.inconsistencias.length;
  if (gravidade === 'ERRO') envelope.ok = false;
}

function core_v2RotinasSanitizeError_(err) {
  var text = err && err.message ? err.message : String(err || 'Erro desconhecido.');
  return text
    .replace(/[A-Za-z0-9_-]{30,}/g, '[id_interno]')
    .replace(/Disponiveis: .+$/i, 'Consulte o Registry para conferir as keys disponiveis.');
}

function core_v2RotinasReadKey_(key) {
  try {
    var sheet = core_getSheetByKey_(key);
    var lastCol = sheet.getLastColumn();
    var headers = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(header) {
          return core_v2RotinasText_(header);
        })
      : [];
    return {
      ok: true,
      key: key,
      sheet: sheet,
      headers: headers,
      records: core_readRecordsByKey_(key, { skipBlankRows: true })
    };
  } catch (err) {
    return {
      ok: false,
      key: key,
      sheet: null,
      headers: [],
      records: [],
      error: core_v2RotinasSanitizeError_(err)
    };
  }
}

function core_v2RotinasReadGroup_(keyMap) {
  var out = {};
  Object.keys(keyMap || {}).forEach(function(name) {
    out[name] = core_v2RotinasReadKey_(keyMap[name]);
  });
  return out;
}

function core_v2RotinasDiagnosticForGroup_(modulo, fluxo, keyMap, options) {
  var opts = core_v2RotinasOptions_(options);
  var data = core_v2RotinasReadGroup_(keyMap);
  var datasets = [];
  var ok = true;

  Object.keys(data).forEach(function(name) {
    var item = data[name];
    if (!item.ok) ok = false;
    datasets.push({
      nome: name,
      key: item.key,
      ok: item.ok,
      totalRegistros: item.records.length,
      totalCabecalhos: item.headers.length,
      cabecalhos: opts.includeRecords ? item.headers : undefined,
      erro: item.error || ''
    });
  });

  return {
    ok: ok,
    modulo: modulo,
    fluxo: fluxo,
    ambiente: core_v2RotinasSafeCurrentEnv_(),
    datasets: datasets,
    resumo: {
      totalDatasets: datasets.length,
      totalIndisponiveis: datasets.filter(function(item) { return !item.ok; }).length
    }
  };
}

function core_v2RotinasSafeCurrentEnv_() {
  try {
    return core_getCurrentEnv_();
  } catch (err) {
    return 'INDEFINIDO';
  }
}

function core_v2RotinasTryStatus_(kind, modulo, fluxo, payload) {
  try {
    if (kind === 'execution') {
      return core_moduleStatusMarkExecution_(modulo, fluxo, payload.capability || '', {
        modeRead: payload.modeRead || '',
        obs: payload.obs || ''
      });
    }
    if (kind === 'success') {
      return core_moduleStatusMarkSuccess_(modulo, fluxo, payload.capability || '', {
        modeRead: payload.modeRead || '',
        obs: payload.obs || ''
      });
    }
    if (kind === 'error') {
      return core_moduleStatusMarkError_(modulo, fluxo, payload.error || payload.message || '', payload.capability || '', {
        modeRead: payload.modeRead || '',
        obs: payload.obs || ''
      });
    }
    if (kind === 'blocked') {
      return core_moduleStatusMarkBlocked_(modulo, fluxo, payload.reasonCode || 'BLOQUEADO', payload.reasonMessage || '', payload.capability || '', payload.modeRead || '', {
        obs: payload.obs || ''
      });
    }
  } catch (err) {
    return {
      ok: false,
      statusDisponivel: false,
      erro: core_v2RotinasSanitizeError_(err)
    };
  }
  return null;
}

function core_v2RotinasPrepareExecution_(modulo, fluxo, options, capability) {
  var opts = core_v2RotinasOptions_(options);
  var result = {
    ok: true,
    opts: opts,
    config: null,
    configDisponivel: false,
    bloqueado: false,
    status: {}
  };

  try {
    result.config = core_getModuleConfig_(modulo, fluxo, {
      ambiente: opts.ambiente || undefined,
      defaultWhenMissing: null
    });
    result.configDisponivel = !!result.config;
  } catch (err) {
    result.configErro = core_v2RotinasSanitizeError_(err);
  }

  if (result.config) {
    var decision = core_modulesConfigEvaluateExecution_(result.config, capability || '', {
      executionType: 'MANUAL'
    });
    result.decision = decision;
    if (result.config.mode === 'DRY_RUN') result.opts.dryRun = true;
    if (!decision.allowed) {
      result.ok = false;
      result.bloqueado = true;
      result.status.blocked = core_v2RotinasTryStatus_('blocked', modulo, fluxo, {
        reasonCode: 'MODULOS_CONFIG',
        reasonMessage: decision.reason,
        capability: capability || '',
        modeRead: result.config.mode || ''
      });
    }
  }

  if (!result.bloqueado) {
    result.status.execution = core_v2RotinasTryStatus_('execution', modulo, fluxo, {
      capability: capability || '',
      modeRead: result.config ? result.config.mode : ''
    });
  }

  return result;
}

function core_v2RotinasFinishSuccess_(execution, modulo, fluxo, capability, obs) {
  execution.status.success = core_v2RotinasTryStatus_('success', modulo, fluxo, {
    capability: capability || '',
    modeRead: execution.config ? execution.config.mode : '',
    obs: obs || ''
  });
}

function core_v2RotinasFinishError_(execution, modulo, fluxo, capability, err) {
  execution.status.error = core_v2RotinasTryStatus_('error', modulo, fluxo, {
    error: err,
    capability: capability || '',
    modeRead: execution.config ? execution.config.mode : ''
  });
}

function core_pessoasV2Diagnostico_(options) {
  var fluxo = CORE_V2_ROTINAS_FLUXOS.PESSOAS_CONFERENCIA;
  return core_v2RotinasDiagnosticForGroup_(fluxo.modulo, fluxo.fluxo, CORE_V2_ROTINAS_KEYS.PESSOAS, options || {});
}

function core_vigenciasV2Diagnostico_(options) {
  var fluxo = CORE_V2_ROTINAS_FLUXOS.VIGENCIAS_CONFERENCIA;
  return core_v2RotinasDiagnosticForGroup_(fluxo.modulo, fluxo.fluxo, CORE_V2_ROTINAS_KEYS.VIGENCIAS, options || {});
}

function core_v2DiagnosticoGeral_(options) {
  options = options || {};
  var pessoas = core_pessoasV2Diagnostico_(options);
  var vigencias = core_vigenciasV2Diagnostico_(options);
  return {
    ok: pessoas.ok && vigencias.ok,
    modulo: 'CORE',
    fluxo: 'DIAGNOSTICO_V2_GERAL',
    pessoas: pessoas,
    vigencias: vigencias,
    resumo: {
      totalDatasets: pessoas.resumo.totalDatasets + vigencias.resumo.totalDatasets,
      totalIndisponiveis: pessoas.resumo.totalIndisponiveis + vigencias.resumo.totalIndisponiveis
    }
  };
}

function core_pessoasV2ConferirConsistencia_(options) {
  var fluxo = CORE_V2_ROTINAS_FLUXOS.PESSOAS_CONFERENCIA;
  var execution = core_v2RotinasPrepareExecution_(fluxo.modulo, fluxo.fluxo, options || {}, '');
  var envelope = core_v2RotinasNewConsistencyEnvelope_(fluxo.modulo, fluxo.fluxo);
  envelope.config = {
    disponivel: execution.configDisponivel,
    bloqueado: execution.bloqueado,
    erro: execution.configErro || ''
  };
  envelope.status = execution.status;

  if (execution.bloqueado) {
    core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'FLUXO', fluxo.fluxo, 'MODULOS_CONFIG', '', 'FLUXO_ATIVO', 'Fluxo bloqueado por MODULOS_CONFIG.', 'Ajustar MODULOS_CONFIG para PESSOAS / CONFERENCIA_V2.');
    return envelope;
  }

  try {
    var data = core_v2RotinasReadGroup_(CORE_V2_ROTINAS_KEYS.PESSOAS);
    core_v2RotinasAddUnavailableDataIssues_(envelope, data, 'PESSOAS');
    if (!envelope.ok) {
      core_v2RotinasFinishError_(execution, fluxo.modulo, fluxo.fluxo, '', 'Datasets V2 indisponiveis.');
      envelope.status = execution.status;
      return envelope;
    }

    core_pessoasV2ConferirData_(envelope, data);
    core_v2RotinasFinishSuccess_(execution, fluxo.modulo, fluxo.fluxo, '', 'Conferencia Pessoas v2 finalizada.');
    envelope.status = execution.status;
    return envelope;
  } catch (err) {
    core_v2RotinasFinishError_(execution, fluxo.modulo, fluxo.fluxo, '', err);
    core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'FLUXO', fluxo.fluxo, 'EXECUCAO', '', 'ERRO_CONTROLADO', core_v2RotinasSanitizeError_(err), 'Revisar erro controlado antes de nova execucao.');
    envelope.status = execution.status;
    return envelope;
  }
}

function core_vigenciasV2ConferirConsistencia_(options) {
  var fluxo = CORE_V2_ROTINAS_FLUXOS.VIGENCIAS_CONFERENCIA;
  var execution = core_v2RotinasPrepareExecution_(fluxo.modulo, fluxo.fluxo, options || {}, '');
  var envelope = core_v2RotinasNewConsistencyEnvelope_(fluxo.modulo, fluxo.fluxo);
  envelope.config = {
    disponivel: execution.configDisponivel,
    bloqueado: execution.bloqueado,
    erro: execution.configErro || ''
  };
  envelope.status = execution.status;

  if (execution.bloqueado) {
    core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'FLUXO', fluxo.fluxo, 'MODULOS_CONFIG', '', 'FLUXO_ATIVO', 'Fluxo bloqueado por MODULOS_CONFIG.', 'Ajustar MODULOS_CONFIG para VIGENCIAS / CONFERENCIA_V2.');
    return envelope;
  }

  try {
    var data = core_v2RotinasReadGroup_(CORE_V2_ROTINAS_KEYS.VIGENCIAS);
    var pessoas = core_v2RotinasReadGroup_({ BASE: CORE_V2_ROTINAS_KEYS.PESSOAS.BASE });
    data.PESSOAS_BASE = pessoas.BASE;
    core_v2RotinasAddUnavailableDataIssues_(envelope, data, 'VIGENCIAS');
    if (!envelope.ok) {
      core_v2RotinasFinishError_(execution, fluxo.modulo, fluxo.fluxo, '', 'Datasets V2 indisponiveis.');
      envelope.status = execution.status;
      return envelope;
    }

    core_vigenciasV2ConferirData_(envelope, data);
    core_v2RotinasFinishSuccess_(execution, fluxo.modulo, fluxo.fluxo, '', 'Conferencia Vigencias v2 finalizada.');
    envelope.status = execution.status;
    return envelope;
  } catch (err) {
    core_v2RotinasFinishError_(execution, fluxo.modulo, fluxo.fluxo, '', err);
    core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'FLUXO', fluxo.fluxo, 'EXECUCAO', '', 'ERRO_CONTROLADO', core_v2RotinasSanitizeError_(err), 'Revisar erro controlado antes de nova execucao.');
    envelope.status = execution.status;
    return envelope;
  }
}

function core_v2RotinasAddUnavailableDataIssues_(envelope, data, entidade) {
  Object.keys(data || {}).forEach(function(name) {
    var item = data[name];
    if (item.ok) return;
    core_v2RotinasAddInconsistency_(
      envelope,
      'ERRO',
      entidade,
      item.key,
      'REGISTRY',
      item.key,
      'KEY_V2_DISPONIVEL',
      'Nao foi possivel ler a key V2 pelo Registry: ' + name + '.',
      item.error || 'Conferir Registry e ambiente.'
    );
  });
}

function core_pessoasV2ConferirData_(envelope, data) {
  var base = data.BASE.records || [];
  var identificadores = data.IDENTIFICADORES.records || [];
  var detalhes = data.MEMBROS_DETALHES.records || [];
  var vinculos = data.VINCULOS.records || [];
  var eventos = data.EVENTOS.records || [];
  envelope.totalVerificado = base.length + identificadores.length + detalhes.length + vinculos.length + eventos.length;

  var pessoasById = core_v2RotinasIndexManyBy_(base, 'ID_PESSOA');
  var detalhesById = core_v2RotinasIndexManyBy_(detalhes, 'ID_PESSOA');
  var vinculosById = core_v2RotinasIndexManyBy_(vinculos, 'ID_PESSOA');
  var vinculosByVinculoId = core_v2RotinasIndexFirstBy_(vinculos, 'ID_VINCULO');
  var activeEmailOwners = {};
  var rgaOwners = {};

  base.forEach(function(pessoa) {
    var idPessoa = core_v2RotinasText_(pessoa.ID_PESSOA);
    var email = core_v2RotinasEmail_(pessoa.EMAIL_PRINCIPAL);
    var cpfDigits = core_v2RotinasDigits_(pessoa.CPF);
    var active = core_v2RotinasPessoaAtiva_(pessoa);

    if (!idPessoa) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'PESSOA', 'linha ' + pessoa.__rowNumber, 'ID_PESSOA', '', 'ID_PESSOA_OBRIGATORIO', 'Pessoa sem ID_PESSOA.', 'Preencher ID_PESSOA tecnico antes de consumir a v2.');
    }
    if (email && !core_isValidEmail_(email)) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'PESSOA', idPessoa, 'EMAIL_PRINCIPAL', pessoa.EMAIL_PRINCIPAL, 'EMAIL_VALIDO', 'EMAIL_PRINCIPAL invalido.', 'Corrigir e-mail principal ou mover para observacao.');
    }
    if (email && active) {
      if (!activeEmailOwners[email]) activeEmailOwners[email] = [];
      activeEmailOwners[email].push(idPessoa || 'linha ' + pessoa.__rowNumber);
    }
    if (cpfDigits && cpfDigits.length !== 11) {
      core_v2RotinasAddInconsistency_(envelope, 'ALERTA', 'PESSOA', idPessoa, 'CPF', pessoa.CPF, 'CPF_11_DIGITOS', 'CPF preenchido fora do padrao de 11 digitos.', 'Padronizar CPF ou deixar vazio quando nao houver dado confiavel.');
    }
  });

  Object.keys(pessoasById).forEach(function(idPessoa) {
    if (pessoasById[idPessoa].length > 1) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'PESSOA', idPessoa, 'ID_PESSOA', idPessoa, 'ID_PESSOA_UNICO', 'ID_PESSOA duplicado em PESSOAS_BASE.', 'Unificar ou corrigir registros duplicados.');
    }
  });

  Object.keys(activeEmailOwners).forEach(function(email) {
    var owners = core_v2RotinasUnique_(activeEmailOwners[email]);
    if (owners.length > 1) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'PESSOA', owners.join('; '), 'EMAIL_PRINCIPAL', email, 'EMAIL_PRINCIPAL_UNICO_ATIVOS', 'EMAIL_PRINCIPAL duplicado em pessoas ativas diferentes.', 'Manter um unico e-mail principal por pessoa ativa.');
    }
  });

  identificadores.forEach(function(record) {
    if (core_v2RotinasStatus_(record.TIPO_IDENTIFICADOR) !== 'RGA') return;
    var rga = core_v2RotinasText_(record.VALOR_IDENTIFICADOR);
    if (!rga) return;
    if (!rgaOwners[rga]) rgaOwners[rga] = [];
    rgaOwners[rga].push(core_v2RotinasText_(record.ID_PESSOA) || 'linha ' + record.__rowNumber);
  });

  Object.keys(rgaOwners).forEach(function(rga) {
    var owners = core_v2RotinasUnique_(rgaOwners[rga]);
    if (owners.length > 1 || rgaOwners[rga].length > 1) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'PESSOA', owners.join('; '), 'RGA', rga, 'RGA_UNICO_IDENTIFICADORES', 'RGA duplicado em PESSOAS_IDENTIFICADORES.', 'Corrigir identificadores duplicados antes de usar como chave de busca.');
    }
  });

  vinculos.forEach(function(vinculo) {
    var idPessoa = core_v2RotinasText_(vinculo.ID_PESSOA);
    var active = core_v2RotinasVinculoAtivo_(vinculo);
    if (active && (!idPessoa || !pessoasById[idPessoa])) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'VINCULO', core_v2RotinasText_(vinculo.ID_VINCULO), 'ID_PESSOA', idPessoa, 'VINCULO_ATIVO_COM_PESSOA_VALIDA', 'Vinculo ativo sem pessoa valida.', 'Corrigir ID_PESSOA ou encerrar o vinculo invalido.');
    }
    if (active && idPessoa && pessoasById[idPessoa] && !core_v2RotinasPessoaAtiva_(pessoasById[idPessoa][0])) {
      core_v2RotinasAddInconsistency_(envelope, 'ALERTA', 'VINCULO', core_v2RotinasText_(vinculo.ID_VINCULO), 'ID_PESSOA', idPessoa, 'PESSOA_ATIVA_COM_VINCULO_ATIVO', 'Pessoa inativa com vinculo ativo.', 'Revisar status cadastral ou encerrar vinculo.');
    }
    if (active && core_v2RotinasTipoVinculo_(vinculo.TIPO_VINCULO) === 'MEMBRO_EFETIVO' && !core_v2RotinasPessoaHasRga_(idPessoa, detalhesById, identificadores)) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'PESSOA', idPessoa, 'RGA', '', 'MEMBRO_EFETIVO_COM_RGA', 'Membro efetivo sem RGA.', 'Registrar RGA em MEMBROS_DETALHES ou PESSOAS_IDENTIFICADORES.');
    }
  });

  eventos.forEach(function(evento) {
    var idPessoa = core_v2RotinasText_(evento.ID_PESSOA);
    var idVinculo = core_v2RotinasText_(evento.ID_VINCULO);
    var rga = core_v2RotinasText_(evento.RGA);
    if (idPessoa && !pessoasById[idPessoa]) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'VINCULO', core_v2RotinasText_(evento.ID_EVENTO_MEMBRO), 'ID_PESSOA', idPessoa, 'EVENTO_COM_PESSOA_OU_VINCULO_VALIDO', 'Evento de vinculo aponta para pessoa inexistente.', 'Corrigir ID_PESSOA do evento ou sanear evento legado.');
    }
    if (idVinculo && !vinculosByVinculoId[idVinculo]) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'VINCULO', core_v2RotinasText_(evento.ID_EVENTO_MEMBRO), 'ID_VINCULO', idVinculo, 'EVENTO_COM_VINCULO_VALIDO', 'Evento de vinculo aponta para vinculo inexistente.', 'Corrigir ID_VINCULO ou deixar vazio se a relacao nao existir na base.');
    }
    if (!idPessoa && rga && !core_v2RotinasFindPessoaIdByRga_(rga, detalhes, identificadores)) {
      core_v2RotinasAddInconsistency_(envelope, 'ALERTA', 'VINCULO', core_v2RotinasText_(evento.ID_EVENTO_MEMBRO), 'RGA', rga, 'EVENTO_COM_PESSOA_RESOLVIDA', 'Evento sem pessoa correspondente por RGA.', 'Resolver ID_PESSOA antes de processar evento.');
    }
  });
}

function core_vigenciasV2ConferirData_(envelope, data) {
  var semestres = data.SEMESTRES.records || [];
  var periodos = data.PERIODOS.records || [];
  var diretorias = data.DIRETORIAS.records || [];
  var cargos = data.CARGOS.records || [];
  var funcoes = data.FUNCOES.records || [];
  var pessoas = data.PESSOAS_BASE.records || [];
  envelope.totalVerificado = semestres.length + periodos.length + diretorias.length + cargos.length + funcoes.length;

  var today = new Date();
  var pessoasById = core_v2RotinasIndexFirstBy_(pessoas, 'ID_PESSOA');
  var cargosByKey = core_v2RotinasIndexFirstBy_(cargos, 'CARGO_KEY');
  var activeSemestres = semestres.filter(function(record) {
    return core_v2RotinasTemporalAtivo_(record, 'STATUS', today);
  });
  var activePeriodos = periodos.filter(function(record) {
    return core_v2RotinasTemporalAtivo_(record, 'STATUS', today);
  });
  var activeFuncoes = funcoes.filter(core_v2RotinasFuncaoVigente_);

  if (!activeSemestres.length) {
    core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'VIGENCIA', 'SEMESTRES', 'STATUS', '', 'SEMESTRE_ATIVO_UNICO', 'Semestre ativo ausente.', 'Marcar exatamente um semestre ativo ou ajustar datas/status.');
  }
  if (activeSemestres.length > 1) {
    core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'VIGENCIA', activeSemestres.map(function(r) { return r.ID_SEMESTRE || r.__rowNumber; }).join('; '), 'STATUS', 'ATIVO', 'SEMESTRE_ATIVO_UNICO', 'Mais de um semestre ativo simultaneo.', 'Manter somente um semestre ativo.');
  }
  if (!activePeriodos.length) {
    core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'VIGENCIA', 'PERIODOS', 'STATUS', '', 'PERIODO_ATIVO_PRESENTE', 'Periodo ativo ausente.', 'Marcar um periodo ativo ou ajustar datas/status.');
  }

  funcoes.forEach(function(funcao) {
    var idPessoa = core_v2RotinasText_(funcao.ID_PESSOA);
    var idVigencia = core_v2RotinasText_(funcao.ID_VIGENCIA) || 'linha ' + funcao.__rowNumber;
    var cargoKey = core_v2RotinasText_(funcao.CARGO_KEY);
    var cargo = cargoKey ? cargosByKey[cargoKey] : null;
    var inicio = core_v2RotinasDate_(funcao.DATA_INICIO);
    var fimPrevista = core_v2RotinasDate_(funcao.DATA_FIM_PREVISTA);
    var fimReal = core_v2RotinasDate_(funcao.DATA_FIM_REAL);
    var fim = fimReal || fimPrevista;

    if (core_v2RotinasFuncaoVigente_(funcao) && !idPessoa) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'VIGENCIA', idVigencia, 'ID_PESSOA', '', 'FUNCAO_VIGENTE_COM_ID_PESSOA', 'Funcao vigente sem ID_PESSOA.', 'Preencher ID_PESSOA antes de gerar resumo atual.');
    }
    if (core_v2RotinasFuncaoVigente_(funcao) && idPessoa && !pessoasById[idPessoa]) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'VIGENCIA', idVigencia, 'ID_PESSOA', idPessoa, 'FUNCAO_VIGENTE_COM_PESSOA_EXISTENTE', 'Funcao vigente com ID_PESSOA inexistente.', 'Corrigir ID_PESSOA ou incluir pessoa na base Pessoas v2.');
    }
    if (inicio && fim && fim < inicio) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'VIGENCIA', idVigencia, 'DATA_FIM', fim, 'DATA_FIM_MAIOR_OU_IGUAL_INICIO', 'Data de fim anterior a data de inicio.', 'Corrigir datas de vigencia.');
    }
    if (!cargo) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'CARGO', idVigencia, 'CARGO_KEY', cargoKey, 'CARGO_KEY_EM_CARGOS_CONFIG', 'Cargo/funcao sem correspondencia em CARGOS_CONFIG.', 'Cadastrar CARGO_KEY em CARGOS_CONFIG ou corrigir a funcao.');
    }
  });

  core_vigenciasV2ConferirCargosExclusivos_(envelope, activeFuncoes, cargosByKey);
  core_vigenciasV2ConferirDiretorias_(envelope, diretorias, activeFuncoes, cargosByKey, today);
}

function core_vigenciasV2ConferirCargosExclusivos_(envelope, activeFuncoes, cargosByKey) {
  for (var i = 0; i < activeFuncoes.length; i++) {
    var a = activeFuncoes[i];
    var cargoA = cargosByKey[core_v2RotinasText_(a.CARGO_KEY)] || {};
    if (!core_v2RotinasIsSim_(cargoA.E_CARGO_UNICO)) continue;
    for (var j = i + 1; j < activeFuncoes.length; j++) {
      var b = activeFuncoes[j];
      if (core_v2RotinasText_(a.CARGO_KEY) !== core_v2RotinasText_(b.CARGO_KEY)) continue;
      if (core_v2RotinasText_(a.ID_DIRETORIA) !== core_v2RotinasText_(b.ID_DIRETORIA)) continue;
      if (!core_v2RotinasIntervalsOverlap_(
        core_v2RotinasDate_(a.DATA_INICIO),
        core_v2RotinasDate_(a.DATA_FIM_REAL) || core_v2RotinasDate_(a.DATA_FIM_PREVISTA),
        core_v2RotinasDate_(b.DATA_INICIO),
        core_v2RotinasDate_(b.DATA_FIM_REAL) || core_v2RotinasDate_(b.DATA_FIM_PREVISTA)
      )) continue;
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'CARGO', core_v2RotinasText_(a.CARGO_KEY), 'CARGO_KEY', core_v2RotinasText_(a.CARGO_KEY), 'CARGO_EXCLUSIVO_SEM_DUPLICIDADE_INTERVALO', 'Cargo exclusivo duplicado no mesmo intervalo.', 'Encerrar ou corrigir uma das vigencias sobrepostas.');
    }
  }
}

function core_vigenciasV2ConferirDiretorias_(envelope, diretorias, activeFuncoes, cargosByKey, today) {
  diretorias.filter(function(record) {
    return core_v2RotinasTemporalAtivo_(record, 'STATUS_DIRETORIA', today);
  }).forEach(function(diretoria) {
    var idDiretoria = core_v2RotinasText_(diretoria.ID_DIRETORIA);
    if (!idDiretoria) return;
    var funcoesDiretoria = activeFuncoes.filter(function(funcao) {
      return core_v2RotinasText_(funcao.ID_DIRETORIA) === idDiretoria;
    });
    var hasPresidente = false;
    var hasVice = false;
    funcoesDiretoria.forEach(function(funcao) {
      var cargo = cargosByKey[core_v2RotinasText_(funcao.CARGO_KEY)] || {};
      var label = core_v2RotinasNormalize_([
        funcao.CARGO_KEY,
        funcao.CARGO_NOME_SNAPSHOT,
        cargo.CARGO_NOME,
        cargo.NOME_PUBLICO
      ].join(' '));
      if (label.indexOf('PRESIDENTE') >= 0 && label.indexOf('VICE') < 0) hasPresidente = true;
      if (label.indexOf('VICE') >= 0 && label.indexOf('PRESIDENTE') >= 0) hasVice = true;
    });
    if (!hasPresidente) {
      core_v2RotinasAddInconsistency_(envelope, 'ERRO', 'VIGENCIA', idDiretoria, 'CARGO_KEY', 'PRESIDENTE', 'DIRETORIA_COM_PRESIDENTE_E_VICE', 'Diretoria vigente sem presidente.', 'Registrar vigencia vigente de presidente para a diretoria.');
    }
    if (!hasVice) {
      core_v2RotinasAddInconsistency_(envelope, 'ALERTA', 'VIGENCIA', idDiretoria, 'CARGO_KEY', 'VICE_PRESIDENTE', 'DIRETORIA_COM_PRESIDENTE_E_VICE', 'Diretoria vigente sem vice quando aplicavel.', 'Registrar vice-presidente ou justificar a excecao normativa.');
    }
  });
}

function core_v2RotinasPessoaAtiva_(pessoa) {
  if (!pessoa) return false;
  if (Object.prototype.hasOwnProperty.call(pessoa, 'ATIVO')) return core_v2RotinasIsSim_(pessoa.ATIVO);
  var status = core_v2RotinasStatus_(pessoa.STATUS_CADASTRAL);
  return status !== 'INATIVO' && status !== 'INATIVA' && status !== 'DESLIGADO' && status !== 'DESLIGADA';
}

function core_v2RotinasVinculoAtivo_(vinculo) {
  return core_v2RotinasIsRecordActive_(vinculo, 'STATUS_VINCULO', 'ATIVO');
}

function core_v2RotinasFuncaoVigente_(funcao) {
  if (core_v2RotinasText_(funcao.DATA_FIM_REAL)) return false;
  return core_v2RotinasIsRecordActive_(funcao, 'STATUS_VIGENCIA', 'ATIVO');
}

function core_v2RotinasTemporalAtivo_(record, statusField, today) {
  var status = core_v2RotinasStatus_(record[statusField || 'STATUS']);
  if (status === 'ATIVO' || status === 'ATIVA' || status === 'VIGENTE') return true;
  if (status === 'INATIVO' || status === 'INATIVA' || status === 'ENCERRADO' || status === 'ENCERRADA') return false;
  var start = core_v2RotinasDate_(record.DATA_INICIO);
  var end = core_v2RotinasDate_(record.DATA_FIM_REAL) || core_v2RotinasDate_(record.DATA_FIM) || core_v2RotinasDate_(record.DATA_FIM_PREVISTA);
  return !!(start && today >= start && (!end || today <= end));
}

function core_v2RotinasPessoaHasRga_(idPessoa, detalhesById, identificadores) {
  if (!idPessoa) return false;
  var detalhes = detalhesById[idPessoa] || [];
  if (detalhes.some(function(record) { return !!core_v2RotinasText_(record.RGA); })) return true;
  return (identificadores || []).some(function(record) {
    return core_v2RotinasText_(record.ID_PESSOA) === idPessoa &&
      core_v2RotinasStatus_(record.TIPO_IDENTIFICADOR) === 'RGA' &&
      core_v2RotinasText_(record.VALOR_IDENTIFICADOR);
  });
}

function core_v2RotinasFindPessoaIdByRga_(rga, detalhes, identificadores) {
  var target = core_v2RotinasText_(rga);
  if (!target) return '';
  for (var i = 0; i < (detalhes || []).length; i++) {
    if (core_v2RotinasText_(detalhes[i].RGA) === target) return core_v2RotinasText_(detalhes[i].ID_PESSOA);
  }
  for (var j = 0; j < (identificadores || []).length; j++) {
    if (core_v2RotinasStatus_(identificadores[j].TIPO_IDENTIFICADOR) === 'RGA' &&
        core_v2RotinasText_(identificadores[j].VALOR_IDENTIFICADOR) === target) {
      return core_v2RotinasText_(identificadores[j].ID_PESSOA);
    }
  }
  return '';
}

function core_v2RotinasUnique_(values) {
  var seen = {};
  var out = [];
  (values || []).forEach(function(value) {
    var key = core_v2RotinasText_(value);
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(key);
  });
  return out;
}

function core_pessoasV2AtualizarResumoOperacional_(options) {
  var fluxo = CORE_V2_ROTINAS_FLUXOS.PESSOAS_ATUALIZACAO;
  var execution = core_v2RotinasPrepareExecution_(fluxo.modulo, fluxo.fluxo, options || {}, 'SYNC');
  if (execution.bloqueado) {
    return core_v2RotinasBlockedUpdateEnvelope_(fluxo.modulo, fluxo.fluxo, execution);
  }

  try {
    var data = core_v2RotinasReadGroup_(CORE_V2_ROTINAS_KEYS.PESSOAS);
    var vigenciasResumo = core_v2RotinasReadKey_(CORE_V2_ROTINAS_KEYS.VIGENCIAS.RESUMO);
    var unavailable = core_v2RotinasUnavailableKeys_(data);
    if (unavailable.length) {
      throw new Error('Keys V2 indisponiveis: ' + unavailable.join(', '));
    }

    var consistencia = core_pessoasV2ConferirConsistencia_({ dryRun: true });
    var rows = core_pessoasV2BuildResumoRows_(data, vigenciasResumo.ok ? vigenciasResumo.records : []);
    var result = core_v2RotinasUpdateEnvelope_(fluxo.modulo, fluxo.fluxo, execution, rows, consistencia);

    if (!execution.opts.dryRun && consistencia.ok === false && !execution.opts.allowWriteWithErrors) {
      result.ok = false;
      result.bloqueado = true;
      result.mensagem = 'Escrita bloqueada por inconsistencias de Pessoas v2. Use allowWriteWithErrors apenas com decisao operacional explicita.';
      core_v2RotinasFinishError_(execution, fluxo.modulo, fluxo.fluxo, 'SYNC', result.mensagem);
      result.status = execution.status;
      return result;
    }

    if (!execution.opts.dryRun) {
      result.escrita = core_v2RotinasWriteWithLock_(function() {
        var target = core_v2RotinasReadKey_(CORE_V2_ROTINAS_KEYS.PESSOAS.RESUMO);
        if (!target.ok) throw new Error(target.error);
        return core_v2RotinasUpsertRows_(target.sheet, rows, {
          primaryKey: 'ID_PESSOA',
          requiredHeaders: CORE_V2_ROTINAS_PESSOAS_RESUMO_HEADERS,
          dryRun: false
        });
      });
    }

    core_v2RotinasFinishSuccess_(execution, fluxo.modulo, fluxo.fluxo, 'SYNC', 'Resumo operacional de Pessoas v2 calculado.');
    result.status = execution.status;
    return result;
  } catch (err) {
    core_v2RotinasFinishError_(execution, fluxo.modulo, fluxo.fluxo, 'SYNC', err);
    return core_v2RotinasErrorUpdateEnvelope_(fluxo.modulo, fluxo.fluxo, execution, err);
  }
}

function core_vigenciasV2AtualizarResumoAtual_(options) {
  var fluxo = CORE_V2_ROTINAS_FLUXOS.VIGENCIAS_ATUALIZACAO;
  var execution = core_v2RotinasPrepareExecution_(fluxo.modulo, fluxo.fluxo, options || {}, 'SYNC');
  if (execution.bloqueado) {
    return core_v2RotinasBlockedUpdateEnvelope_(fluxo.modulo, fluxo.fluxo, execution);
  }

  try {
    var data = core_v2RotinasReadGroup_(CORE_V2_ROTINAS_KEYS.VIGENCIAS);
    var pessoasData = core_v2RotinasReadGroup_({
      BASE: CORE_V2_ROTINAS_KEYS.PESSOAS.BASE,
      MEMBROS_DETALHES: CORE_V2_ROTINAS_KEYS.PESSOAS.MEMBROS_DETALHES
    });
    var unavailable = core_v2RotinasUnavailableKeys_(data).concat(core_v2RotinasUnavailableKeys_(pessoasData));
    if (unavailable.length) {
      throw new Error('Keys V2 indisponiveis: ' + unavailable.join(', '));
    }

    var consistencia = core_vigenciasV2ConferirConsistencia_({ dryRun: true });
    var rows = core_vigenciasV2BuildResumoRows_(data, pessoasData);
    var result = core_v2RotinasUpdateEnvelope_(fluxo.modulo, fluxo.fluxo, execution, rows, consistencia);

    if (!execution.opts.dryRun && consistencia.ok === false && !execution.opts.allowWriteWithErrors) {
      result.ok = false;
      result.bloqueado = true;
      result.mensagem = 'Escrita bloqueada por inconsistencias de Vigencias v2. Use allowWriteWithErrors apenas com decisao operacional explicita.';
      core_v2RotinasFinishError_(execution, fluxo.modulo, fluxo.fluxo, 'SYNC', result.mensagem);
      result.status = execution.status;
      return result;
    }

    if (!execution.opts.dryRun) {
      result.escrita = core_v2RotinasWriteWithLock_(function() {
        var target = core_v2RotinasReadKey_(CORE_V2_ROTINAS_KEYS.VIGENCIAS.RESUMO);
        if (!target.ok) throw new Error(target.error);
        return core_v2RotinasUpsertRows_(target.sheet, rows, {
          primaryKey: 'ID_VIGENCIA',
          fallbackKey: 'ID_PESSOA',
          requiredHeaders: CORE_V2_ROTINAS_VIGENCIAS_RESUMO_HEADERS,
          dryRun: false
        });
      });
    }

    core_v2RotinasFinishSuccess_(execution, fluxo.modulo, fluxo.fluxo, 'SYNC', 'Resumo atual de Vigencias v2 calculado.');
    result.status = execution.status;
    return result;
  } catch (err) {
    core_v2RotinasFinishError_(execution, fluxo.modulo, fluxo.fluxo, 'SYNC', err);
    return core_v2RotinasErrorUpdateEnvelope_(fluxo.modulo, fluxo.fluxo, execution, err);
  }
}

function core_pessoasV2BuildResumoRows_(data, vigenciasResumoRecords) {
  var base = data.BASE.records || [];
  var detalhesById = core_v2RotinasIndexFirstBy_(data.MEMBROS_DETALHES.records || [], 'ID_PESSOA');
  var vinculosById = core_v2RotinasIndexManyBy_(data.VINCULOS.records || [], 'ID_PESSOA');
  var identificadores = data.IDENTIFICADORES.records || [];
  var vigResumoById = core_v2RotinasIndexFirstBy_(vigenciasResumoRecords || [], 'ID_PESSOA');
  var now = new Date();

  return base.map(function(pessoa) {
    var idPessoa = core_v2RotinasText_(pessoa.ID_PESSOA);
    var detalhes = detalhesById[idPessoa] || {};
    var vinculo = core_v2RotinasPickCurrentVinculo_(vinculosById[idPessoa] || []) || {};
    var vigResumo = vigResumoById[idPessoa] || {};
    var tipo = core_v2RotinasTipoVinculo_(vinculo.TIPO_VINCULO);
    var active = core_v2RotinasVinculoAtivo_(vinculo);
    var perfilBase = tipo === 'MEMBRO_EFETIVO' && active ? 'MEMBRO' : '';
    if (tipo === 'MEMBRO_EM_ESPERA' && active) perfilBase = 'MEMBRO_EM_ESPERA';
    if (tipo === 'EGRESSO') perfilBase = 'EGRESSO';
    var perfilVigencia = vigResumo.PERFIL_PORTAL_GERADO || vigResumo.PERFIS_PORTAL_CALCULADOS || '';
    var portalAtivo = active && tipo !== 'EGRESSO' ? 'SIM' : 'NAO';

    return {
      ID_PESSOA: idPessoa,
      NOME_EXIBICAO: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || '',
      EMAIL_PRINCIPAL: pessoa.EMAIL_PRINCIPAL || '',
      EMAIL: pessoa.EMAIL_PRINCIPAL || '',
      RGA: core_v2RotinasGetRga_(idPessoa, detalhes, identificadores),
      CPF: pessoa.CPF || '',
      TIPO_VINCULO_ATUAL: vinculo.TIPO_VINCULO || '',
      STATUS_VINCULO_ATUAL: vinculo.STATUS_VINCULO || '',
      DATA_INICIO_VINCULO: vinculo.DATA_INICIO || '',
      DATA_FIM_VINCULO: vinculo.DATA_FIM || '',
      PORTAL_ATIVO: portalAtivo,
      PERFIL_PORTAL_BASE: perfilBase,
      PERFIL_PORTAL_CALCULADO: core_v2RotinasJoinDistinct_([perfilBase, perfilVigencia]),
      CARGO_FUNCAO_ATUAL: vigResumo.OCUPACAO || vigResumo.CARGO_FUNCAO_ATUAL || '',
      ULTIMA_ATUALIZACAO: now,
      OBS_RESUMO: idPessoa ? '' : 'SEM_ID_PESSOA'
    };
  });
}

function core_vigenciasV2BuildResumoRows_(data, pessoasData) {
  var pessoasById = core_v2RotinasIndexFirstBy_(pessoasData.BASE.records || [], 'ID_PESSOA');
  var detalhesById = core_v2RotinasIndexFirstBy_(pessoasData.MEMBROS_DETALHES.records || [], 'ID_PESSOA');
  var cargosByKey = core_v2RotinasIndexFirstBy_(data.CARGOS.records || [], 'CARGO_KEY');
  var now = new Date();

  return (data.FUNCOES.records || []).filter(core_v2RotinasFuncaoVigente_).map(function(funcao) {
    var idPessoa = core_v2RotinasText_(funcao.ID_PESSOA);
    var pessoa = pessoasById[idPessoa] || {};
    var detalhes = detalhesById[idPessoa] || {};
    var cargoKey = core_v2RotinasText_(funcao.CARGO_KEY);
    var cargo = cargosByKey[cargoKey] || {};
    var ocupacao = funcao.CARGO_NOME_SNAPSHOT || cargo.NOME_PUBLICO || cargo.CARGO_NOME || cargoKey;
    var permissoes = core_v2RotinasPermissionsFromCargo_(cargo);
    var perfil = core_v2RotinasIsSim_(cargo.GERA_PERFIL_PORTAL) ? (cargo.PERFIL_PORTAL_PADRAO || cargo.CARGO_KEY || '') : '';
    var grupo = cargo.GRUPO_FUNCAO || cargo.GRUPO_CARGO || '';

    return {
      ID_VIGENCIA: core_v2RotinasText_(funcao.ID_VIGENCIA) || core_v2RotinasBuildSyntheticVigenciaId_(funcao),
      ID_PESSOA: idPessoa,
      NOME_EXIBICAO: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || '',
      RGA: detalhes.RGA || '',
      OCUPACAO: ocupacao,
      GRUPO_FUNCAO: grupo,
      DATA_INICIO: funcao.DATA_INICIO || '',
      DATA_FIM_PREVISTA: funcao.DATA_FIM_PREVISTA || '',
      DATA_FIM_REAL: funcao.DATA_FIM_REAL || '',
      STATUS_VIGENCIA: funcao.STATUS_VIGENCIA || '',
      PERFIL_PORTAL_GERADO: perfil,
      PERMISSOES_GERADAS: permissoes.join('; '),
      CARGO_ATUAL_VISIVEL: ocupacao,
      APARECE_DIRETORIA_PUBLICA: core_v2RotinasStatus_(funcao.TIPO_FUNCAO) === 'DIRETORIA' ? 'SIM' : 'NAO',
      CARGO_FUNCAO_ATUAL: ocupacao,
      TIPO_FUNCAO_ATUAL: funcao.TIPO_FUNCAO || cargo.TIPO_FUNCAO || '',
      GRUPO_FUNCAO_ATUAL: grupo,
      ID_DIRETORIA_ATUAL: funcao.ID_DIRETORIA || '',
      PERFIS_PORTAL_CALCULADOS: perfil,
      PERMISSOES_CALCULADAS: permissoes.join('; '),
      DATA_INICIO_FUNCAO_ATUAL: funcao.DATA_INICIO || '',
      ULTIMA_ATUALIZACAO: now
    };
  }).sort(function(a, b) {
    return String(a.NOME_EXIBICAO || a.ID_PESSOA).localeCompare(String(b.NOME_EXIBICAO || b.ID_PESSOA));
  });
}

function core_v2RotinasPickCurrentVinculo_(vinculos) {
  var priority = {
    MEMBRO_EFETIVO_ATIVO: 1,
    MEMBRO_EM_ESPERA_ATIVO: 2,
    EGRESSO: 3,
    OUTRO_ATIVO: 4,
    OUTRO: 9
  };
  return (vinculos || []).slice().sort(function(a, b) {
    var pa = core_v2RotinasVinculoRank_(a, priority);
    var pb = core_v2RotinasVinculoRank_(b, priority);
    if (pa !== pb) return pa - pb;
    return (core_v2RotinasDate_(b.DATA_INICIO) || 0) - (core_v2RotinasDate_(a.DATA_INICIO) || 0);
  })[0] || null;
}

function core_v2RotinasVinculoRank_(vinculo, priority) {
  var tipo = core_v2RotinasTipoVinculo_(vinculo.TIPO_VINCULO);
  var active = core_v2RotinasVinculoAtivo_(vinculo);
  if (tipo === 'MEMBRO_EFETIVO' && active) return priority.MEMBRO_EFETIVO_ATIVO;
  if (tipo === 'MEMBRO_EM_ESPERA' && active) return priority.MEMBRO_EM_ESPERA_ATIVO;
  if (tipo === 'EGRESSO') return priority.EGRESSO;
  if (active) return priority.OUTRO_ATIVO;
  return priority.OUTRO;
}

function core_v2RotinasGetRga_(idPessoa, detalhes, identificadores) {
  var fallback = detalhes && detalhes.RGA ? detalhes.RGA : '';
  for (var i = (identificadores || []).length - 1; i >= 0; i--) {
    var record = identificadores[i];
    if (core_v2RotinasText_(record.ID_PESSOA) !== idPessoa) continue;
    if (core_v2RotinasStatus_(record.TIPO_IDENTIFICADOR) !== 'RGA') continue;
    if (core_v2RotinasText_(record.VALOR_IDENTIFICADOR)) return record.VALOR_IDENTIFICADOR;
  }
  return fallback || '';
}

function core_v2RotinasJoinDistinct_(values) {
  return core_v2RotinasUnique_(values).join('; ');
}

function core_v2RotinasPermissionsFromCargo_(cargo) {
  var permissions = [];
  [
    'PODE_VER_AREA_DIRETORIA',
    'PODE_GERENCIAR_ATIVIDADES',
    'PODE_REGISTRAR_CHAMADA',
    'PODE_EDITAR_ATIVIDADE',
    'PODE_ANALISAR_JUSTIFICATIVAS',
    'PODE_GERENCIAR_CERTIFICADOS',
    'PODE_GERENCIAR_COMUNICACAO'
  ].forEach(function(field) {
    if (core_v2RotinasIsSim_(cargo[field])) permissions.push(field);
  });
  return permissions;
}

function core_v2RotinasBuildSyntheticVigenciaId_(funcao) {
  return [
    core_v2RotinasText_(funcao.ID_PESSOA),
    core_v2RotinasText_(funcao.CARGO_KEY),
    core_v2RotinasText_(funcao.ID_DIRETORIA),
    core_v2RotinasText_(funcao.DATA_INICIO)
  ].filter(String).join('|') || 'linha ' + funcao.__rowNumber;
}

function core_v2RotinasUnavailableKeys_(data) {
  return Object.keys(data || {}).filter(function(name) {
    return !data[name].ok;
  }).map(function(name) {
    return data[name].key;
  });
}

function core_v2RotinasUpdateEnvelope_(modulo, fluxo, execution, rows, consistencia) {
  return {
    ok: true,
    modulo: modulo,
    fluxo: fluxo,
    dryRun: execution.opts.dryRun,
    config: {
      disponivel: execution.configDisponivel,
      modo: execution.config ? execution.config.mode : '',
      erro: execution.configErro || ''
    },
    totalCalculado: rows.length,
    totalEscreveria: execution.opts.dryRun ? rows.length : 0,
    totalEscrito: 0,
    previa: rows.slice(0, execution.opts.limit),
    consistencia: {
      ok: consistencia.ok,
      totalVerificado: consistencia.totalVerificado,
      totalInconsistencias: consistencia.totalInconsistencias,
      amostraInconsistencias: (consistencia.inconsistencias || []).slice(0, execution.opts.limit)
    },
    status: execution.status
  };
}

function core_v2RotinasBlockedUpdateEnvelope_(modulo, fluxo, execution) {
  return {
    ok: false,
    modulo: modulo,
    fluxo: fluxo,
    dryRun: true,
    bloqueado: true,
    mensagem: 'Fluxo bloqueado por MODULOS_CONFIG.',
    config: {
      disponivel: execution.configDisponivel,
      modo: execution.config ? execution.config.mode : '',
      motivo: execution.decision ? execution.decision.reason : '',
      erro: execution.configErro || ''
    },
    status: execution.status
  };
}

function core_v2RotinasErrorUpdateEnvelope_(modulo, fluxo, execution, err) {
  return {
    ok: false,
    modulo: modulo,
    fluxo: fluxo,
    dryRun: execution.opts.dryRun,
    erro: core_v2RotinasSanitizeError_(err),
    status: execution.status
  };
}

function core_v2RotinasWriteWithLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function core_v2RotinasEnsureHeaders_(sheet, requiredHeaders, dryRun) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(header) {
        return core_v2RotinasText_(header);
      })
    : [];
  var headerMap = core_buildHeaderIndexMap_(headers, {
    normalize: true,
    oneBased: false,
    keepFirst: true
  });
  var missing = [];
  (requiredHeaders || []).forEach(function(header) {
    if (!Object.prototype.hasOwnProperty.call(headerMap, core_normalizeHeader_(header))) missing.push(header);
  });

  if (missing.length && dryRun !== true) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }

  return {
    headers: headers,
    missingHeaders: missing,
    addedHeaders: dryRun === true ? [] : missing
  };
}

function core_v2RotinasUpsertRows_(sheet, rows, opts) {
  opts = opts || {};
  var headerResult = core_v2RotinasEnsureHeaders_(sheet, opts.requiredHeaders || [], opts.dryRun === true);
  var headers = headerResult.headers;
  var records = core_readSheetRecords_(sheet, { skipBlankRows: true });
  var primaryKey = opts.primaryKey;
  var fallbackKey = opts.fallbackKey || '';
  var byPrimary = {};
  var byFallback = {};

  records.forEach(function(record) {
    var primary = core_v2RotinasText_(record[primaryKey]);
    if (primary && !byPrimary[primary]) byPrimary[primary] = record;
    if (fallbackKey) {
      var fallback = core_v2RotinasText_(record[fallbackKey]);
      if (fallback && !byFallback[fallback]) byFallback[fallback] = record;
    }
  });

  var updated = 0;
  var appended = 0;
  (rows || []).forEach(function(row) {
    var primary = core_v2RotinasText_(row[primaryKey]);
    var existing = primary ? byPrimary[primary] : null;
    if (!existing && fallbackKey) {
      var fallback = core_v2RotinasText_(row[fallbackKey]);
      existing = fallback ? byFallback[fallback] : null;
    }
    if (existing && existing.__rowNumber) {
      var merged = core_v2RotinasClone_(existing);
      Object.keys(row).forEach(function(key) {
        merged[key] = row[key];
      });
      sheet.getRange(existing.__rowNumber, 1, 1, headers.length).setValues([core_buildRowFromObjectByHeaders_(headers, merged)]);
      updated++;
    } else {
      core_appendObjectByHeaders_(sheet, row, { headerRow: 1 });
      appended++;
    }
  });

  return {
    ok: true,
    updated: updated,
    appended: appended,
    totalWritten: updated + appended,
    addedHeaders: headerResult.addedHeaders,
    missingHeadersBeforeWrite: headerResult.missingHeaders
  };
}

function coreV2RunTesteDiagnosticoGeral_() {
  return core_v2DiagnosticoGeral_({ limit: 5 });
}

function coreV2RunTestePessoasResumo_() {
  return core_pessoasV2AtualizarResumoOperacional_({
    dryRun: true,
    limit: 5
  });
}

function coreV2RunTesteVigenciasResumo_() {
  return core_vigenciasV2AtualizarResumoAtual_({
    dryRun: true,
    limit: 5
  });
}
