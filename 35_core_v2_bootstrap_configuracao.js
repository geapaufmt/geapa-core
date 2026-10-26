/**
 * ============================================================
 * 35_core_v2_bootstrap_configuracao.js
 * ============================================================
 *
 * Conferencia/bootstrap seguro das configuracoes operacionais V2.
 *
 * Regras:
 * - dryRun e o padrao;
 * - cria somente linhas ausentes, quando explicitamente solicitado;
 * - nunca sobrescreve, apaga ou corrige duplicidades automaticamente;
 * - nunca escreve fora de DEV;
 * - usa cabecalhos e LockService para escrita.
 */

var CORE_V2_BOOTSTRAP_FLUXO = Object.freeze({
  modulo: 'CORE',
  fluxo: 'BOOTSTRAP_CONFIGURACAO_V2'
});

var CORE_V2_BOOTSTRAP_REQUIRED_FLOWS = Object.freeze([
  Object.freeze({ modulo: 'CORE', fluxo: 'JOB_DIARIO_V2', sync: true }),
  Object.freeze({ modulo: 'PESSOAS', fluxo: 'ATUALIZACAO_V2', sync: true }),
  Object.freeze({ modulo: 'PESSOAS', fluxo: 'CONFERENCIA_V2', sync: false }),
  Object.freeze({ modulo: 'VIGENCIAS', fluxo: 'ATUALIZACAO_V2', sync: true }),
  Object.freeze({ modulo: 'VIGENCIAS', fluxo: 'CONFERENCIA_V2', sync: false }),
  Object.freeze({ modulo: 'ATIVIDADES', fluxo: 'ATUALIZACAO_PORTAL_V2', sync: true }),
  Object.freeze({ modulo: 'ATIVIDADES', fluxo: 'CONFERENCIA_V2', sync: false }),
  Object.freeze({ modulo: 'ATIVIDADES', fluxo: 'FREQUENCIA_V2', sync: true })
]);

function core_v2BootstrapOptions_(options) {
  options = options || {};
  var ambiente = core_v2BootstrapNormalizeEnv_(options.ambiente || 'DEV');
  return {
    dryRun: options.dryRun !== false,
    criarAusentes: options.criarAusentes === true,
    ambiente: ambiente,
    ambienteValido: ambiente === 'DEV' || ambiente === 'PROD',
    origem: String(options.origem || 'MANUAL').trim() || 'MANUAL'
  };
}

function core_v2BootstrapNormalize_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  }).replace(/\s+/g, '_');
}

function core_v2BootstrapNormalizeEnv_(value) {
  return core_v2BootstrapNormalize_(value || 'DEV');
}

function core_v2BootstrapBuildKey_(modulo, fluxo, ambiente) {
  return [
    core_v2BootstrapNormalize_(modulo),
    core_v2BootstrapNormalize_(fluxo || 'GERAL'),
    core_v2BootstrapNormalizeEnv_(ambiente || 'DEV')
  ].join('|');
}

function core_v2BootstrapBuildStatusKey_(modulo, fluxo) {
  return [
    core_v2BootstrapNormalize_(modulo),
    core_v2BootstrapNormalize_(fluxo || 'GERAL')
  ].join('|');
}

function core_v2BootstrapNewReport_(opts) {
  return {
    ok: true,
    modulo: CORE_V2_BOOTSTRAP_FLUXO.modulo,
    fluxo: CORE_V2_BOOTSTRAP_FLUXO.fluxo,
    dryRun: opts.dryRun,
    criarAusentes: opts.criarAusentes,
    ambiente: opts.ambiente,
    ambienteValido: opts.ambienteValido,
    origem: opts.origem,
    prontoHomologacao: false,
    bloqueadoEscrita: false,
    motivoBloqueioEscrita: '',
    config: core_v2BootstrapNewSection_('MODULOS_CONFIG'),
    status: core_v2BootstrapNewSection_('MODULOS_STATUS'),
    criados: {
      config: [],
      status: []
    },
    totais: {},
    avisos: [],
    erros: []
  };
}

function core_v2BootstrapNewSection_(sheetName) {
  return {
    sheetName: sheetName,
    ok: true,
    existente: [],
    ausente: [],
    duplicado: [],
    invalido: [],
    seriaCriado: [],
    criado: [],
    erros: []
  };
}

function coreV2_conferirConfiguracao_(options) {
  var opts = core_v2BootstrapOptions_(options || {});
  opts.dryRun = true;
  opts.criarAusentes = false;
  return core_v2BootstrapBuildReport_(opts);
}

function coreV2_bootstrapConfiguracao_(options) {
  var opts = core_v2BootstrapOptions_(options || {});
  var report = core_v2BootstrapBuildReport_(opts);

  if (!opts.criarAusentes || opts.dryRun) return report;

  if (!opts.ambienteValido || opts.ambiente !== 'DEV') {
    report.ok = false;
    report.prontoHomologacao = false;
    report.bloqueadoEscrita = true;
    report.motivoBloqueioEscrita = opts.ambienteValido
      ? 'Bootstrap V2 nunca escreve fora de DEV.'
      : 'Ambiente invalido para bootstrap V2.';
    report.erros.push(report.motivoBloqueioEscrita);
    return report;
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    report.ok = false;
    report.prontoHomologacao = false;
    report.bloqueadoEscrita = true;
    report.motivoBloqueioEscrita = 'LOCK_INDISPONIVEL';
    report.erros.push('Nao foi possivel obter lock para bootstrap de configuracao V2.');
    return report;
  }

  try {
    var locked = core_v2BootstrapBuildReport_(opts);
    var createdConfig = core_v2BootstrapCreateMissingConfig_(locked, opts);
    var createdStatus = core_v2BootstrapCreateMissingStatus_(locked, opts);
    core_modulesConfigCacheClear_();

    var after = core_v2BootstrapBuildReport_(opts);
    after.criados.config = createdConfig;
    after.criados.status = createdStatus;
    after.config.criado = createdConfig;
    after.status.criado = createdStatus;
    after.totais.configCriadas = createdConfig.length;
    after.totais.statusCriados = createdStatus.length;
    return after;
  } catch (err) {
    report.ok = false;
    report.prontoHomologacao = false;
    report.erros.push(core_v2RotinasSanitizeError_(err));
    return report;
  } finally {
    lock.releaseLock();
  }
}

function core_v2BootstrapBuildReport_(opts) {
  var report = core_v2BootstrapNewReport_(opts);

  core_v2BootstrapInspectConfig_(report, opts);
  core_v2BootstrapInspectStatus_(report, opts);
  core_v2BootstrapFinalizeReport_(report);
  return report;
}

function core_v2BootstrapInspectConfig_(report, opts) {
  try {
    var sheet = core_getModulesConfigSheet_();
    var data = core_v2BootstrapReadRows_(sheet);
    var missingHeaders = core_v2BootstrapMissingHeaders_(data.headerMap, [
      'MODULO',
      'FLUXO',
      'ATIVO',
      'MODO',
      'AMBIENTE',
      'PERMITE_TRIGGER',
      'PERMITE_EMAIL',
      'PERMITE_INBOX',
      'PERMITE_SYNC',
      'PERMITE_DRIVE'
    ]);

    if (missingHeaders.length) {
      report.config.ok = false;
      report.config.erros.push('Cabecalhos ausentes: ' + missingHeaders.join(', '));
      return;
    }

    var index = core_v2BootstrapIndexConfigRows_(data);
    report.config.invalido = index.invalidos;
    report.config.duplicado = index.duplicados;

    CORE_V2_BOOTSTRAP_REQUIRED_FLOWS.forEach(function(def) {
      var key = core_v2BootstrapBuildKey_(def.modulo, def.fluxo, opts.ambiente);
      var rows = index.byKey[key] || [];
      var item = core_v2BootstrapFlowItem_(def, opts.ambiente);
      if (rows.length === 1) {
        item.lineNo = rows[0].lineNo;
        item.mode = rows[0].mode || '';
        item.active = rows[0].active || '';
        report.config.existente.push(item);
      } else if (rows.length === 0) {
        report.config.ausente.push(item);
        if (opts.dryRun || !opts.criarAusentes) report.config.seriaCriado.push(item);
      }
    });
  } catch (err) {
    report.config.ok = false;
    report.config.erros.push(core_v2RotinasSanitizeError_(err));
  }
}

function core_v2BootstrapInspectStatus_(report, opts) {
  try {
    var sheet = core_getModulesStatusSheet_();
    var data = core_v2BootstrapReadRows_(sheet);
    var missingHeaders = core_v2BootstrapMissingHeaders_(data.headerMap, CORE_MODULES_STATUS_HEADERS);

    if (missingHeaders.length) {
      report.status.ok = false;
      report.status.erros.push('Cabecalhos ausentes: ' + missingHeaders.join(', '));
      return;
    }

    var index = core_v2BootstrapIndexStatusRows_(data);
    report.status.invalido = index.invalidos;
    report.status.duplicado = index.duplicados;

    CORE_V2_BOOTSTRAP_REQUIRED_FLOWS.forEach(function(def) {
      var key = core_v2BootstrapBuildStatusKey_(def.modulo, def.fluxo);
      var rows = index.byKey[key] || [];
      var item = core_v2BootstrapFlowItem_(def, '');
      if (rows.length === 1) {
        item.lineNo = rows[0].lineNo;
        report.status.existente.push(item);
      } else if (rows.length === 0) {
        report.status.ausente.push(item);
        if (opts.dryRun || !opts.criarAusentes) report.status.seriaCriado.push(item);
      }
    });
  } catch (err) {
    report.status.ok = false;
    report.status.erros.push(core_v2RotinasSanitizeError_(err));
  }
}

function core_v2BootstrapReadRows_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(header) {
        return String(header || '').trim();
      })
    : [];
  var headerMap = core_buildHeaderIndexMap_(headers, {
    normalize: true,
    oneBased: true,
    keepFirst: true
  });
  var lastRow = sheet.getLastRow();
  var values = lastRow >= 2 && lastCol > 0
    ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
    : [];
  return {
    sheet: sheet,
    headers: headers,
    headerMap: headerMap,
    values: values
  };
}

function core_v2BootstrapMissingHeaders_(headerMap, required) {
  return (required || []).filter(function(headerName) {
    return !core_getCol_(headerMap, headerName);
  });
}

function core_v2BootstrapCell_(row, headerMap, headerName) {
  var col = core_getCol_(headerMap, headerName);
  if (!col) return '';
  return row[col - 1];
}

function core_v2BootstrapIndexConfigRows_(data) {
  var byKey = {};
  var invalidos = [];
  data.values.forEach(function(row, idx) {
    var lineNo = idx + 2;
    var moduleName = core_v2BootstrapNormalize_(core_v2BootstrapCell_(row, data.headerMap, 'MODULO'));
    var flowName = core_v2BootstrapNormalize_(core_v2BootstrapCell_(row, data.headerMap, 'FLUXO') || 'GERAL');
    var ambiente = core_v2BootstrapNormalize_(core_v2BootstrapCell_(row, data.headerMap, 'AMBIENTE') || 'PROD');
    if (!moduleName && !flowName) return;
    if (!moduleName) {
      invalidos.push({ lineNo: lineNo, motivo: 'MODULO_AUSENTE' });
      return;
    }

    var issues = [];
    if (ambiente !== 'DEV' && ambiente !== 'PROD') issues.push('AMBIENTE_INVALIDO');
    if (['ON', 'OFF', 'MANUAL', 'DRY_RUN'].indexOf(core_v2BootstrapNormalize_(core_v2BootstrapCell_(row, data.headerMap, 'MODO') || 'ON')) < 0) {
      issues.push('MODO_INVALIDO');
    }
    ['ATIVO', 'PERMITE_TRIGGER', 'PERMITE_EMAIL', 'PERMITE_INBOX', 'PERMITE_SYNC', 'PERMITE_DRIVE'].forEach(function(headerName) {
      var value = core_v2BootstrapNormalize_(core_v2BootstrapCell_(row, data.headerMap, headerName));
      if (value !== 'SIM' && value !== 'NAO') issues.push(headerName + '_INVALIDO');
    });
    if (issues.length) {
      invalidos.push({ lineNo: lineNo, key: moduleName + ' / ' + flowName + ' / ' + ambiente, motivo: issues.join(', ') });
    }

    var key = core_v2BootstrapBuildKey_(moduleName, flowName, ambiente);
    if (!byKey[key]) byKey[key] = [];
    byKey[key].push({
      lineNo: lineNo,
      moduleName: moduleName,
      flowName: flowName,
      ambiente: ambiente,
      mode: core_v2BootstrapNormalize_(core_v2BootstrapCell_(row, data.headerMap, 'MODO') || ''),
      active: core_v2BootstrapNormalize_(core_v2BootstrapCell_(row, data.headerMap, 'ATIVO') || '')
    });
  });
  return {
    byKey: byKey,
    invalidos: invalidos,
    duplicados: core_v2BootstrapBuildDuplicateList_(byKey, true)
  };
}

function core_v2BootstrapIndexStatusRows_(data) {
  var byKey = {};
  var invalidos = [];
  data.values.forEach(function(row, idx) {
    var lineNo = idx + 2;
    var moduleName = core_v2BootstrapNormalize_(core_v2BootstrapCell_(row, data.headerMap, 'MODULO'));
    var flowName = core_v2BootstrapNormalize_(core_v2BootstrapCell_(row, data.headerMap, 'FLUXO') || 'GERAL');
    if (!moduleName && !flowName) return;
    if (!moduleName) {
      invalidos.push({ lineNo: lineNo, motivo: 'MODULO_AUSENTE' });
      return;
    }
    var key = core_v2BootstrapBuildStatusKey_(moduleName, flowName);
    if (!byKey[key]) byKey[key] = [];
    byKey[key].push({
      lineNo: lineNo,
      moduleName: moduleName,
      flowName: flowName
    });
  });
  return {
    byKey: byKey,
    invalidos: invalidos,
    duplicados: core_v2BootstrapBuildDuplicateList_(byKey, false)
  };
}

function core_v2BootstrapBuildDuplicateList_(byKey, includeAmbiente) {
  var out = [];
  Object.keys(byKey || {}).forEach(function(key) {
    var rows = byKey[key] || [];
    if (rows.length > 1) {
      var first = rows[0] || {};
      out.push({
        modulo: first.moduleName || '',
        fluxo: first.flowName || '',
        ambiente: includeAmbiente ? first.ambiente || '' : '',
        linhas: rows.map(function(row) { return row.lineNo; })
      });
    }
  });
  return out;
}

function core_v2BootstrapFlowItem_(def, ambiente) {
  return {
    modulo: def.modulo,
    fluxo: def.fluxo,
    ambiente: ambiente || ''
  };
}

function core_v2BootstrapCreateMissingConfig_(report, opts) {
  if (report.config.erros.length || report.config.duplicado.length || report.config.invalido.length) return [];
  var sheet = core_getModulesConfigSheet_();
  var created = [];
  report.config.ausente.forEach(function(item) {
    var def = core_v2BootstrapFindRequiredFlow_(item.modulo, item.fluxo);
    var payload = core_v2BootstrapBuildConfigPayload_(def, opts);
    core_appendObjectByHeaders_(sheet, payload, { headerRow: 1 });
    created.push(item);
  });
  return created;
}

function core_v2BootstrapCreateMissingStatus_(report, opts) {
  if (report.status.erros.length || report.status.duplicado.length || report.status.invalido.length) return [];
  var sheet = core_getModulesStatusSheet_();
  var created = [];
  report.status.ausente.forEach(function(item) {
    var payload = core_v2BootstrapBuildStatusPayload_(item, opts);
    core_appendObjectByHeaders_(sheet, payload, { headerRow: 1 });
    created.push(item);
  });
  return created;
}

function core_v2BootstrapFindRequiredFlow_(modulo, fluxo) {
  var moduleKey = core_v2BootstrapNormalize_(modulo);
  var flowKey = core_v2BootstrapNormalize_(fluxo);
  for (var i = 0; i < CORE_V2_BOOTSTRAP_REQUIRED_FLOWS.length; i++) {
    var def = CORE_V2_BOOTSTRAP_REQUIRED_FLOWS[i];
    if (def.modulo === moduleKey && def.fluxo === flowKey) return def;
  }
  return { modulo: moduleKey, fluxo: flowKey, sync: false };
}

function core_v2BootstrapBuildConfigPayload_(def, opts) {
  return {
    MODULO: def.modulo,
    FLUXO: def.fluxo,
    ATIVO: 'SIM',
    MODO: 'DRY_RUN',
    AMBIENTE: opts.ambiente,
    PERMITE_TRIGGER: 'NAO',
    PERMITE_EMAIL: 'NAO',
    PERMITE_INBOX: 'NAO',
    PERMITE_SYNC: def.sync ? 'SIM' : 'NAO',
    PERMITE_DRIVE: 'NAO',
    JANELA_MINUTOS: '',
    ULTIMA_ALTERACAO: new Date(),
    ALTERADO_POR: opts.origem,
    OBS: 'Criado por bootstrap V2 seguro. Revisar antes de liberar escrita real ou trigger.'
  };
}

function core_v2BootstrapBuildStatusPayload_(item, opts) {
  return {
    MODULO: item.modulo,
    FLUXO: item.fluxo,
    EXECUCOES_24H: 0,
    BLOQUEIOS_24H: 0,
    SUCESSOS_24H: 0,
    ERROS_24H: 0,
    OBS: 'Criado por bootstrap V2 seguro em ' + opts.ambiente + '.'
  };
}

function core_v2BootstrapFinalizeReport_(report) {
  if (!report.ambienteValido) {
    report.erros.push('Ambiente invalido para configuracao V2: ' + (report.ambiente || '(vazio)') + '. Use DEV ou PROD.');
  }

  report.config.ok = report.config.ok &&
    report.config.ausente.length === 0 &&
    report.config.duplicado.length === 0 &&
    report.config.invalido.length === 0 &&
    report.config.erros.length === 0;
  report.status.ok = report.status.ok &&
    report.status.ausente.length === 0 &&
    report.status.duplicado.length === 0 &&
    report.status.invalido.length === 0 &&
    report.status.erros.length === 0;

  report.totais = {
    configExistentes: report.config.existente.length,
    configAusentes: report.config.ausente.length,
    configDuplicados: report.config.duplicado.length,
    configInvalidos: report.config.invalido.length,
    configSeriamCriadas: report.config.seriaCriado.length,
    statusExistentes: report.status.existente.length,
    statusAusentes: report.status.ausente.length,
    statusDuplicados: report.status.duplicado.length,
    statusInvalidos: report.status.invalido.length,
    statusSeriamCriados: report.status.seriaCriado.length
  };

  if (report.config.duplicado.length || report.status.duplicado.length) {
    report.avisos.push('Existem duplicidades; o bootstrap apenas reporta e nao corrige automaticamente.');
  }
  if (report.config.ausente.length || report.status.ausente.length) {
    report.avisos.push('Existem linhas ausentes para homologacao V2.');
  }
  if (report.config.invalido.length || report.status.invalido.length) {
    report.erros.push('Existem linhas invalidas nas configuracoes/status operacionais.');
  }
  if (report.config.erros.length || report.status.erros.length) {
    report.erros = report.erros.concat(report.config.erros, report.status.erros);
  }

  report.prontoHomologacao = report.config.ok && report.status.ok;
  report.ok = report.erros.length === 0;
  return report;
}

function coreV2_runTesteBootstrapDryRun_() {
  return coreV2_bootstrapConfiguracao_({
    dryRun: true,
    criarAusentes: true,
    ambiente: 'DEV',
    origem: 'TESTE_MANUAL'
  });
}
