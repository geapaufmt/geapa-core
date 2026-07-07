/**
 * ============================================================
 * 34_core_v2_jobs.js
 * ============================================================
 *
 * Orquestracao manual/trigger-safe dos jobs V2.
 *
 * Esta camada nao instala trigger automaticamente e nao depende de
 * Atividades estar carregado no mesmo projeto. Quando o runner de
 * Atividades nao estiver disponivel, a etapa fica marcada como SKIPPED.
 */

var CORE_V2_JOB_DIARIO = Object.freeze({
  modulo: 'CORE',
  fluxo: 'JOB_DIARIO_V2',
  capabilityTrigger: 'TRIGGER',
  handler: 'coreV2_jobDiarioManutencaoTrigger',
  hour: 4
});

function core_v2JobOptions_(options, config) {
  options = options || {};
  var executionType = core_v2RotinasNormalize_(options.executionType || 'MANUAL');
  var configMode = config ? core_v2RotinasNormalize_(config.mode) : '';
  var dryRun;

  if (executionType === 'TRIGGER') {
    dryRun = options.dryRun === true || configMode === 'DRY_RUN' || !config;
  } else {
    dryRun = options.dryRun !== false || configMode === 'DRY_RUN';
  }

  return {
    dryRun: dryRun,
    executionType: executionType,
    limit: Math.max(1, Number(options.limit || 5)),
    ambiente: options.ambiente || '',
    atividadesJob: options.atividadesJob,
    atividadesAtualizarViews: options.atividadesAtualizarViews,
    atividadesConferir: options.atividadesConferir
  };
}

function core_v2JobPrepare_(options) {
  options = options || {};
  var config = null;
  var configErro = '';
  var configDisponivel = false;
  var decision = null;
  var capability = core_v2RotinasNormalize_(options.executionType || 'MANUAL') === 'TRIGGER'
    ? CORE_V2_JOB_DIARIO.capabilityTrigger
    : '';

  try {
    config = core_getModuleConfig_(CORE_V2_JOB_DIARIO.modulo, CORE_V2_JOB_DIARIO.fluxo, {
      ambiente: options.ambiente || undefined,
      defaultWhenMissing: null
    });
    configDisponivel = !!config;
  } catch (err) {
    configErro = core_v2RotinasSanitizeError_(err);
  }

  var opts = core_v2JobOptions_(options, config);
  if (config) {
    decision = core_modulesConfigEvaluateExecution_(config, capability, {
      executionType: opts.executionType
    });
  }

  var blocked = decision && !decision.allowed;
  var status = {};
  if (blocked) {
    status.blocked = core_v2RotinasTryStatus_('blocked', CORE_V2_JOB_DIARIO.modulo, CORE_V2_JOB_DIARIO.fluxo, {
      reasonCode: 'MODULOS_CONFIG',
      reasonMessage: decision.reason,
      capability: capability,
      modeRead: config.mode || ''
    });
  } else {
    status.execution = core_v2RotinasTryStatus_('execution', CORE_V2_JOB_DIARIO.modulo, CORE_V2_JOB_DIARIO.fluxo, {
      capability: capability,
      modeRead: config ? config.mode : '',
      obs: 'Job diario V2 iniciado.'
    });
  }

  return {
    ok: !blocked,
    blocked: blocked,
    opts: opts,
    config: config,
    configDisponivel: configDisponivel,
    configErro: configErro,
    decision: decision,
    capability: capability,
    status: status
  };
}

function coreV2_jobDiarioManutencao_(options) {
  var prepared = core_v2JobPrepare_(options || {});
  var startedAt = new Date();
  var result = {
    ok: true,
    modulo: CORE_V2_JOB_DIARIO.modulo,
    fluxo: CORE_V2_JOB_DIARIO.fluxo,
    dryRun: prepared.opts.dryRun,
    executionType: prepared.opts.executionType,
    startedAt: startedAt,
    finishedAt: '',
    durationMs: 0,
    config: {
      disponivel: prepared.configDisponivel,
      modo: prepared.config ? prepared.config.mode : '',
      bloqueado: prepared.blocked,
      motivo: prepared.decision ? prepared.decision.reason : '',
      erro: prepared.configErro || ''
    },
    steps: [],
    totals: {
      ok: 0,
      warnings: 0,
      errors: 0,
      skipped: 0
    },
    status: prepared.status
  };

  if (prepared.blocked) {
    result.ok = false;
    result.finishedAt = new Date();
    result.durationMs = result.finishedAt.getTime() - startedAt.getTime();
    core_v2JobCountTotals_(result);
    return result;
  }

  try {
    core_v2JobRunStep_(result, 'REGISTRY_DIAGNOSTICO', function() {
      return core_v2DiagnosticoGeral_({ limit: prepared.opts.limit });
    });
    core_v2JobRunStep_(result, 'PESSOAS_ATUALIZAR_RESUMO', function() {
      return core_pessoasV2AtualizarResumoOperacional_({
        dryRun: prepared.opts.dryRun,
        limit: prepared.opts.limit
      });
    });
    core_v2JobRunStep_(result, 'PESSOAS_CONFERIR', function() {
      return core_pessoasV2ConferirConsistencia_({ limit: prepared.opts.limit });
    });
    core_v2JobRunStep_(result, 'VIGENCIAS_ATUALIZAR_RESUMO', function() {
      return core_vigenciasV2AtualizarResumoAtual_({
        dryRun: prepared.opts.dryRun,
        limit: prepared.opts.limit
      });
    });
    core_v2JobRunStep_(result, 'VIGENCIAS_CONFERIR', function() {
      return core_vigenciasV2ConferirConsistencia_({ limit: prepared.opts.limit });
    });
    core_v2JobRunStep_(result, 'ATIVIDADES_ATUALIZAR_VIEWS', function() {
      return core_v2JobRunAtividadesAtualizacao_(prepared.opts);
    });
    core_v2JobRunStep_(result, 'ATIVIDADES_CONFERIR', function() {
      return core_v2JobRunAtividadesConferencia_(prepared.opts);
    });
  } catch (err) {
    result.steps.push(core_v2JobStepError_('JOB_DIARIO_ERRO_CONTROLADO', err));
  }

  result.finishedAt = new Date();
  result.durationMs = result.finishedAt.getTime() - startedAt.getTime();
  core_v2JobCountTotals_(result);
  result.ok = result.totals.errors === 0;
  result.summary = core_v2JobBuildFinalSummary_(result);

  if (result.ok) {
    result.status.success = core_v2RotinasTryStatus_('success', CORE_V2_JOB_DIARIO.modulo, CORE_V2_JOB_DIARIO.fluxo, {
      capability: prepared.capability,
      modeRead: prepared.config ? prepared.config.mode : '',
      obs: result.summary
    });
  } else {
    result.status.error = core_v2RotinasTryStatus_('error', CORE_V2_JOB_DIARIO.modulo, CORE_V2_JOB_DIARIO.fluxo, {
      error: result.summary,
      capability: prepared.capability,
      modeRead: prepared.config ? prepared.config.mode : '',
      obs: result.summary
    });
  }

  return result;
}

function coreV2_jobDiarioManutencaoTrigger() {
  return coreV2_jobDiarioManutencao_({
    executionType: 'TRIGGER'
  });
}

function coreV2_runTesteJobDiarioDryRun_() {
  return coreV2_jobDiarioManutencao_({
    dryRun: true,
    executionType: 'MANUAL',
    limit: 5
  });
}

function core_v2JobRunStep_(jobResult, stepName, fn) {
  var startedAt = new Date();
  try {
    var raw = fn();
    var step = core_v2JobSummarizeStep_(stepName, raw);
    step.durationMs = new Date().getTime() - startedAt.getTime();
    jobResult.steps.push(step);
    return step;
  } catch (err) {
    var errorStep = core_v2JobStepError_(stepName, err);
    errorStep.durationMs = new Date().getTime() - startedAt.getTime();
    jobResult.steps.push(errorStep);
    return errorStep;
  }
}

function core_v2JobSummarizeStep_(stepName, raw) {
  raw = raw || {};
  var skipped = raw.skipped === true;
  var ok = skipped ? true : raw.ok !== false;
  var errors = Number(raw.totalErros || raw.totalErrors || 0);
  var warnings = Number(raw.totalAvisos || raw.totalWarnings || 0);

  if (raw.totalInconsistencias) {
    errors = (raw.inconsistencias || []).filter(function(item) { return item.gravidade === 'ERRO'; }).length;
    warnings += Math.max(0, Number(raw.totalInconsistencias || 0) - errors);
  }
  if (raw.consistencia && raw.consistencia.ok === false) {
    errors += 1;
  }
  if (raw.erro || raw.error) {
    errors += 1;
    ok = false;
  }
  if (errors > 0) ok = false;

  return {
    step: stepName,
    ok: ok,
    status: skipped ? 'SKIPPED' : (ok ? 'OK' : 'ERROR'),
    skipped: skipped,
    dryRun: raw.dryRun === true,
    totalCalculado: Number(raw.totalCalculado || 0),
    totalVerificado: Number(raw.totalVerificado || 0),
    totalInconsistencias: Number(raw.totalInconsistencias || (raw.consistencia && raw.consistencia.totalInconsistencias) || 0),
    errors: errors,
    warnings: warnings,
    message: core_v2JobSafeMessage_(raw.mensagem || raw.message || raw.erro || raw.error || raw.motivo || ''),
    detail: core_v2JobSafeDetail_(raw)
  };
}

function core_v2JobStepError_(stepName, err) {
  return {
    step: stepName,
    ok: false,
    status: 'ERROR',
    skipped: false,
    dryRun: false,
    totalCalculado: 0,
    totalVerificado: 0,
    totalInconsistencias: 0,
    errors: 1,
    warnings: 0,
    message: core_v2JobSafeMessage_(core_v2RotinasSanitizeError_(err)),
    detail: {}
  };
}

function core_v2JobSafeDetail_(raw) {
  var out = {};
  [
    'modulo',
    'fluxo',
    'dryRun',
    'totalCalculado',
    'totalVerificado',
    'totalInconsistencias',
    'totalEscrito',
    'bloqueado'
  ].forEach(function(field) {
    if (Object.prototype.hasOwnProperty.call(raw, field)) out[field] = raw[field];
  });
  if (raw.resumo && raw.resumo.totalIndisponiveis != null) {
    out.totalIndisponiveis = raw.resumo.totalIndisponiveis;
  }
  if (raw.escrita) {
    out.escrita = {
      updated: Number(raw.escrita.updated || 0),
      appended: Number(raw.escrita.appended || 0),
      totalWritten: Number(raw.escrita.totalWritten || 0),
      addedHeaders: (raw.escrita.addedHeaders || []).slice(0, 20)
    };
  }
  if (raw.conferencia) {
    out.conferencia = {
      ok: raw.conferencia.ok !== false,
      totalVerificado: Number(raw.conferencia.totalVerificado || 0),
      totalInconsistencias: Number(raw.conferencia.totalInconsistencias || 0)
    };
  }
  return out;
}

function core_v2JobSafeMessage_(value) {
  return core_v2RotinasSanitizeError_(value || '').replace(/[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}/g, '[email]');
}

function core_v2JobRunAtividadesAtualizacao_(opts) {
  var runner = null;
  if (typeof opts.atividadesAtualizarViews === 'function') runner = opts.atividadesAtualizarViews;
  else if (typeof opts.atividadesJob === 'function') runner = opts.atividadesJob;
  else if (typeof atividadesV2_jobPortal === 'function') runner = atividadesV2_jobPortal;

  if (!runner) {
    return {
      ok: true,
      skipped: true,
      modulo: 'ATIVIDADES',
      fluxo: 'PORTAL_V2',
      motivo: 'Runner atividadesV2_jobPortal nao disponivel neste projeto. Execute pelo geapa-atividades ou passe callback em options.'
    };
  }

  return runner({
    dryRun: opts.dryRun,
    executionType: opts.executionType,
    limit: opts.limit
  });
}

function core_v2JobRunAtividadesConferencia_(opts) {
  var runner = null;
  if (typeof opts.atividadesConferir === 'function') runner = opts.atividadesConferir;
  else if (typeof atividadesV2_conferirPortal === 'function') runner = atividadesV2_conferirPortal;

  if (!runner) {
    return {
      ok: true,
      skipped: true,
      modulo: 'ATIVIDADES',
      fluxo: 'CONFERENCIA_PORTAL_V2',
      motivo: 'Conferencia de Atividades V2 nao disponivel neste projeto. Se atividadesV2_jobPortal ja rodou, confira o retorno no modulo de Atividades.'
    };
  }

  return runner({
    dryRun: true,
    executionType: opts.executionType,
    limit: opts.limit
  });
}

function core_v2JobCountTotals_(result) {
  result.totals = {
    ok: 0,
    warnings: 0,
    errors: 0,
    skipped: 0
  };
  (result.steps || []).forEach(function(step) {
    if (step.skipped) result.totals.skipped++;
    else if (step.ok) result.totals.ok++;
    result.totals.warnings += Number(step.warnings || 0);
    result.totals.errors += Number(step.errors || 0);
  });
  return result.totals;
}

function core_v2JobBuildFinalSummary_(result) {
  return [
    'Job V2 diario',
    result.ok ? 'OK' : 'COM_ERROS',
    'dryRun=' + (result.dryRun ? 'SIM' : 'NAO'),
    'steps=' + (result.steps || []).length,
    'errors=' + result.totals.errors,
    'warnings=' + result.totals.warnings,
    'skipped=' + result.totals.skipped
  ].join(' | ');
}

function coreV2_instalarTriggerJobDiario_(options) {
  options = options || {};
  var hour = Math.max(0, Math.min(23, Number(options.hour == null ? CORE_V2_JOB_DIARIO.hour : options.hour)));
  var existing = coreV2_removerTriggerJobDiario_();
  ScriptApp.newTrigger(CORE_V2_JOB_DIARIO.handler)
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
  return {
    ok: true,
    installed: true,
    deletedBeforeInstall: existing.deleted,
    handler: CORE_V2_JOB_DIARIO.handler,
    schedule: 'Todo dia as ' + hour + 'h',
    note: 'Instalador manual. O handler respeita MODULOS_CONFIG em CORE / JOB_DIARIO_V2.'
  };
}

function coreV2_removerTriggerJobDiario_() {
  var triggers = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === CORE_V2_JOB_DIARIO.handler;
  });
  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  return {
    ok: true,
    deleted: triggers.length,
    handler: CORE_V2_JOB_DIARIO.handler
  };
}

function coreV2_listarTriggerJobDiario_() {
  var triggers = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === CORE_V2_JOB_DIARIO.handler;
  });
  return {
    ok: true,
    handler: CORE_V2_JOB_DIARIO.handler,
    installedCount: triggers.length,
    triggerSource: triggers.length ? String(triggers[0].getTriggerSource()) : '',
    eventType: triggers.length ? String(triggers[0].getEventType()) : '',
    uniqueIds: triggers.map(function(trigger) { return trigger.getUniqueId(); }),
    note: 'A API do Apps Script nao expoe todos os detalhes da agenda configurada.'
  };
}
