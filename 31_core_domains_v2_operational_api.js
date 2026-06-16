/**
 * ============================================================
 * 31_core_domains_v2_operational_api.js
 * ============================================================
 *
 * APIs operacionais permanentes para consumo de Pessoas v2 e Vigencias v2.
 *
 * Regras:
 * - nao executa migracao legado -> v2;
 * - nao altera planilhas legadas;
 * - leituras retornam objetos derivados das bases v2;
 * - recalculos escrevem apenas caches v2 e exigem confirmacao explicita.
 */

function core_domainsV2NewReadReport_(name) {
  return core_domainsV2AuditNewReport_(name || 'DOMINIOS_V2');
}

function core_domainsV2OpenPessoas_(report) {
  return core_domainsV2AuditOpenDomain_('PESSOAS', report || core_domainsV2NewReadReport_('PESSOAS_V2'));
}

function core_domainsV2OpenVigencias_(report) {
  return core_domainsV2AuditOpenDomain_('VIGENCIAS', report || core_domainsV2NewReadReport_('VIGENCIAS_V2'));
}

function core_domainsV2CloneRecord_(record) {
  var out = {};
  Object.keys(record || {}).forEach(function(key) {
    if (key !== '__rowNumber') out[key] = record[key];
  });
  return out;
}

function core_domainsV2Email_(value) {
  return String(value || '').trim().toLowerCase();
}

function core_domainsV2Rga_(value) {
  return String(value || '').trim();
}

function core_domainsV2Active_(record) {
  return core_domainsV2AuditStatus_(record.STATUS_VINCULO || record.STATUS || record.STATUS_VIGENCIA) === 'ATIVO' ||
    core_domainsV2AuditStatus_(record.STATUS_VINCULO || record.STATUS || record.STATUS_VIGENCIA) === 'ATIVA' ||
    core_domainsV2AuditIsSim_(record.ATIVO);
}

function core_domainsV2CurrentVigencia_(record) {
  return core_domainsV2ResumoVigenciaAtual_(record);
}

function core_domainsV2IndexFirstBy_(records, field) {
  var out = {};
  (records || []).forEach(function(record) {
    var key = String(record[field] || '').trim();
    if (key && !out[key]) out[key] = record;
  });
  return out;
}

function core_domainsV2IndexManyBy_(records, field) {
  return core_domainsV2AuditIndexBy_(records || [], field);
}

function core_domainsV2FindPessoaById_(pessoasData, idPessoa) {
  var id = String(idPessoa || '').trim();
  if (!id) return null;
  var records = (pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [];
  for (var i = 0; i < records.length; i++) {
    if (String(records[i].ID_PESSOA || '').trim() === id) return records[i];
  }
  return null;
}

function core_domainsV2FindPessoaIdByIdentifier_(pessoasData, type, value) {
  var normalizedType = core_domainsV2AuditStatus_(type);
  var normalizedValue = normalizedType === 'EMAIL'
    ? core_domainsV2Email_(value)
    : core_domainsV2Rga_(value);
  if (!normalizedValue) return '';

  var identifiers = (pessoasData.PESSOAS_IDENTIFICADORES && pessoasData.PESSOAS_IDENTIFICADORES.records) || [];
  for (var i = identifiers.length - 1; i >= 0; i--) {
    var record = identifiers[i];
    var currentType = core_domainsV2AuditStatus_(record.TIPO_IDENTIFICADOR);
    var currentValue = currentType === 'EMAIL'
      ? core_domainsV2Email_(record.VALOR_IDENTIFICADOR)
      : core_domainsV2Rga_(record.VALOR_IDENTIFICADOR);
    if (currentType === normalizedType && currentValue === normalizedValue && core_domainsV2AuditIsSim_(record.ATIVO)) {
      return String(record.ID_PESSOA || '').trim();
    }
  }
  return '';
}

function core_domainsV2FindPessoaIdByEmail_(pessoasData, email) {
  var normalized = core_domainsV2Email_(email);
  if (!normalized) return '';
  var base = (pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [];
  for (var i = 0; i < base.length; i++) {
    if (core_domainsV2Email_(base[i].EMAIL_PRINCIPAL) === normalized) return String(base[i].ID_PESSOA || '').trim();
  }
  return core_domainsV2FindPessoaIdByIdentifier_(pessoasData, 'EMAIL', normalized);
}

function core_domainsV2FindPessoaIdByRga_(pessoasData, rga) {
  var normalized = core_domainsV2Rga_(rga);
  if (!normalized) return '';
  var detalhes = (pessoasData.MEMBROS_DETALHES && pessoasData.MEMBROS_DETALHES.records) || [];
  for (var i = detalhes.length - 1; i >= 0; i--) {
    if (core_domainsV2Rga_(detalhes[i].RGA) === normalized) return String(detalhes[i].ID_PESSOA || '').trim();
  }
  return core_domainsV2FindPessoaIdByIdentifier_(pessoasData, 'RGA', normalized);
}

function core_domainsV2PessoaBundle_(pessoasData, idPessoa) {
  var pessoa = core_domainsV2FindPessoaById_(pessoasData, idPessoa);
  if (!pessoa) return null;

  var id = String(idPessoa || '').trim();
  var byPessoa = function(sheetName) {
    return ((pessoasData[sheetName] && pessoasData[sheetName].records) || []).filter(function(record) {
      return String(record.ID_PESSOA || '').trim() === id;
    }).map(core_domainsV2CloneRecord_);
  };

  return {
    pessoa: core_domainsV2CloneRecord_(pessoa),
    identificadores: byPessoa('PESSOAS_IDENTIFICADORES'),
    membrosDetalhes: byPessoa('MEMBROS_DETALHES')[0] || null,
    vinculos: byPessoa('VINCULOS_GEAPA'),
    resumoOperacional: byPessoa('PESSOAS_RESUMO_OPERACIONAL')[0] || null,
    comunicacaoConsentimentos: byPessoa('PESSOAS_COMUNICACAO_CONSENTIMENTOS'),
    portalExcecoes: byPessoa('PORTAL_ACESSOS_EXCECOES')
  };
}

function corePessoasGetById_(idPessoa) {
  var report = core_domainsV2NewReadReport_('PESSOAS_GET_BY_ID');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  return core_domainsV2PessoaBundle_(pessoasData, idPessoa);
}

function corePessoasFindByEmail_(email) {
  var report = core_domainsV2NewReadReport_('PESSOAS_FIND_BY_EMAIL');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  var idPessoa = core_domainsV2FindPessoaIdByEmail_(pessoasData, email);
  return idPessoa ? core_domainsV2PessoaBundle_(pessoasData, idPessoa) : null;
}

function corePessoasFindByRga_(rga) {
  var report = core_domainsV2NewReadReport_('PESSOAS_FIND_BY_RGA');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  var idPessoa = core_domainsV2FindPessoaIdByRga_(pessoasData, rga);
  return idPessoa ? core_domainsV2PessoaBundle_(pessoasData, idPessoa) : null;
}

function corePessoasGetOperationalSummary_(idPessoa) {
  var report = core_domainsV2NewReadReport_('PESSOAS_GET_OPERATIONAL_SUMMARY');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  var id = String(idPessoa || '').trim();
  var resumo = (pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [];
  for (var i = 0; i < resumo.length; i++) {
    if (String(resumo[i].ID_PESSOA || '').trim() === id) return core_domainsV2CloneRecord_(resumo[i]);
  }
  return null;
}

function core_domainsV2ListByActiveLinkType_(tipoVinculo, opts) {
  opts = opts || {};
  var report = core_domainsV2NewReadReport_('PESSOAS_LIST_BY_LINK');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));

  var tipo = core_domainsV2AuditTipoVinculo_(tipoVinculo);
  var baseById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [], 'ID_PESSOA');
  var detalhesById = core_domainsV2IndexFirstBy_((pessoasData.MEMBROS_DETALHES && pessoasData.MEMBROS_DETALHES.records) || [], 'ID_PESSOA');
  var resumoById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [], 'ID_PESSOA');
  var out = [];
  var seen = {};

  ((pessoasData.VINCULOS_GEAPA && pessoasData.VINCULOS_GEAPA.records) || []).forEach(function(vinculo) {
    var idPessoa = String(vinculo.ID_PESSOA || '').trim();
    if (!idPessoa || seen[idPessoa]) return;
    if (core_domainsV2AuditTipoVinculo_(vinculo.TIPO_VINCULO) !== tipo) return;
    if (opts.includeInactive !== true && !core_domainsV2Active_(vinculo)) return;
    var pessoa = baseById[idPessoa] || {};
    var detalhes = detalhesById[idPessoa] || {};
    var resumo = resumoById[idPessoa] || {};
    out.push({
      idPessoa: idPessoa,
      nome: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || resumo.NOME_EXIBICAO || '',
      email: pessoa.EMAIL_PRINCIPAL || resumo.EMAIL || '',
      rga: detalhes.RGA || resumo.RGA || '',
      tipoVinculo: vinculo.TIPO_VINCULO || '',
      statusVinculo: vinculo.STATUS_VINCULO || '',
      vinculo: core_domainsV2CloneRecord_(vinculo),
      resumoOperacional: core_domainsV2CloneRecord_(resumo)
    });
    seen[idPessoa] = true;
  });

  return out.sort(function(a, b) {
    return String(a.nome || a.idPessoa).localeCompare(String(b.nome || b.idPessoa));
  });
}

function corePessoasListCurrentMembers_(opts) {
  return core_domainsV2ListByActiveLinkType_('MEMBRO_EFETIVO', opts || {});
}

function corePessoasListExMembers_(opts) {
  return core_domainsV2ListByActiveLinkType_('EGRESSO', opts || {});
}

function corePessoasListWaitingMembers_(opts) {
  return core_domainsV2ListByActiveLinkType_('MEMBRO_EM_ESPERA', opts || {});
}

function core_domainsV2ListDetailWithPessoa_(detailSheetName, opts) {
  opts = opts || {};
  var report = core_domainsV2NewReadReport_('PESSOAS_LIST_DETAIL');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  var baseById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [], 'ID_PESSOA');
  return ((pessoasData[detailSheetName] && pessoasData[detailSheetName].records) || []).filter(function(record) {
    return opts.includeInactive === true || core_domainsV2AuditIsSim_(record.ATIVO);
  }).map(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    var pessoa = baseById[idPessoa] || {};
    return {
      idPessoa: idPessoa,
      nome: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || '',
      email: pessoa.EMAIL_PRINCIPAL || '',
      pessoa: core_domainsV2CloneRecord_(pessoa),
      detalhe: core_domainsV2CloneRecord_(record)
    };
  });
}

function corePessoasListAcademicCollaborators_(opts) {
  return core_domainsV2ListDetailWithPessoa_('COLABORADORES_ACADEMICOS', opts || {});
}

function corePessoasListExternalParticipants_(opts) {
  return core_domainsV2ListDetailWithPessoa_('PARTICIPANTES_EXTERNOS_DETALHES', opts || {});
}

function coreVigenciasListCurrentFunctions_(opts) {
  opts = opts || {};
  var report = core_domainsV2NewReadReport_('VIGENCIAS_LIST_CURRENT_FUNCTIONS');
  var vigenciasData = core_domainsV2OpenVigencias_(report);
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Vigencias v2 indisponivel: ' + JSON.stringify(report.erros));
  var pessoaById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [], 'ID_PESSOA');
  var cargoByKey = core_domainsV2IndexFirstBy_((vigenciasData.CARGOS_CONFIG && vigenciasData.CARGOS_CONFIG.records) || [], 'CARGO_KEY');
  var idFilter = String(opts.idPessoa || '').trim();
  return ((vigenciasData.VIGENCIAS_FUNCOES && vigenciasData.VIGENCIAS_FUNCOES.records) || []).filter(function(record) {
    if (!core_domainsV2CurrentVigencia_(record)) return false;
    if (idFilter && String(record.ID_PESSOA || '').trim() !== idFilter) return false;
    return true;
  }).map(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    var cargo = cargoByKey[String(record.CARGO_KEY || '').trim()] || {};
    var pessoa = pessoaById[idPessoa] || {};
    return {
      idPessoa: idPessoa,
      nome: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || '',
      vigencia: core_domainsV2CloneRecord_(record),
      cargoConfig: core_domainsV2CloneRecord_(cargo)
    };
  });
}

function coreVigenciasGetCurrentFunctionByPessoa_(idPessoa) {
  var list = coreVigenciasListCurrentFunctions_({ idPessoa: idPessoa });
  list.sort(function(a, b) {
    return core_domainsV2ResumoTimestamp_(b.vigencia.DATA_INICIO) - core_domainsV2ResumoTimestamp_(a.vigencia.DATA_INICIO);
  });
  return list[0] || null;
}

function coreVigenciasGetCurrentSummaryByPessoa_(idPessoa) {
  var report = core_domainsV2NewReadReport_('VIGENCIAS_GET_CURRENT_SUMMARY_BY_PESSOA');
  var vigenciasData = core_domainsV2OpenVigencias_(report);
  if (report.totalErros) throw new Error('Vigencias v2 indisponivel: ' + JSON.stringify(report.erros));
  var id = String(idPessoa || '').trim();
  var records = (vigenciasData.VIGENCIAS_RESUMO_ATUAL && vigenciasData.VIGENCIAS_RESUMO_ATUAL.records) || [];
  for (var i = 0; i < records.length; i++) {
    if (String(records[i].ID_PESSOA || '').trim() === id) return core_domainsV2CloneRecord_(records[i]);
  }
  return null;
}

function coreVigenciasGetPortalPermissionsByPessoa_(idPessoa) {
  var functions = coreVigenciasListCurrentFunctions_({ idPessoa: idPessoa });
  var perfis = [];
  var permissoes = [];
  functions.forEach(function(item) {
    var cargo = item.cargoConfig || {};
    if (core_domainsV2AuditIsSim_(cargo.GERA_PERFIL_PORTAL)) {
      core_domainsV2ResumoPushDistinct_(perfis, cargo.PERFIL_PORTAL_PADRAO || cargo.CARGO_KEY);
    }
    core_domainsV2ResumoPermissionsFromCargo_(cargo).forEach(function(permission) {
      core_domainsV2ResumoPushDistinct_(permissoes, permission);
    });
  });
  return {
    idPessoa: String(idPessoa || '').trim(),
    perfisPortalCalculados: perfis,
    permissoesCalculadas: permissoes,
    funcoesConsideradas: functions.length
  };
}

function core_domainsV2PickActiveVinculo_(vinculos) {
  return core_domainsV2PickCurrentVinculo_(vinculos);
}

function core_domainsV2NormalizeTipoVinculo_(value) {
  var tipo = core_domainsV2AuditTipoVinculo_(value);
  if (tipo === 'EX_MEMBRO' || tipo === 'EX-MEMBRO') return 'EGRESSO';
  return tipo;
}

function core_domainsV2PickCurrentVinculo_(vinculos) {
  var priority = {
    MEMBRO_EFETIVO_ATIVO: 1,
    MEMBRO_EM_ESPERA_ATIVO: 2,
    EGRESSO: 3,
    OUTRO_ATIVO: 4,
    MEMBRO_EFETIVO_ENCERRADO: 5,
    MEMBRO_EM_ESPERA_ENCERRADO: 6,
    OUTRO: 9
  };
  return (vinculos || []).slice().sort(function(a, b) {
    var pa = core_domainsV2CurrentVinculoRank_(a, priority);
    var pb = core_domainsV2CurrentVinculoRank_(b, priority);
    if (pa !== pb) return pa - pb;
    return core_domainsV2ResumoTimestamp_(b.DATA_INICIO) - core_domainsV2ResumoTimestamp_(a.DATA_INICIO);
  })[0] || null;
}

function core_domainsV2CurrentVinculoRank_(vinculo, priority) {
  var tipo = core_domainsV2NormalizeTipoVinculo_(vinculo.TIPO_VINCULO);
  var active = core_domainsV2Active_(vinculo);
  if (tipo === 'MEMBRO_EFETIVO' && active) return priority.MEMBRO_EFETIVO_ATIVO;
  if (tipo === 'MEMBRO_EM_ESPERA' && active) return priority.MEMBRO_EM_ESPERA_ATIVO;
  if (tipo === 'EGRESSO') return priority.EGRESSO;
  if (active) return priority.OUTRO_ATIVO;
  if (tipo === 'MEMBRO_EFETIVO') return priority.MEMBRO_EFETIVO_ENCERRADO;
  if (tipo === 'MEMBRO_EM_ESPERA') return priority.MEMBRO_EM_ESPERA_ENCERRADO;
  return priority.OUTRO;
}

function core_domainsV2Date_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  var text = String(value || '').trim();
  var m = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    var year = Number(m[3]);
    if (year < 100) year += 2000;
    return new Date(year, Number(m[2]) - 1, Number(m[1]));
  }
  var parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function core_domainsV2DaysBetween_(start, end) {
  if (!start || !end || end < start) return 0;
  var a = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  var b = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.floor((b.getTime() - a.getTime()) / 86400000) + 1;
}

function core_domainsV2IntervalsOverlap_(aStart, aEnd, bStart, bEnd) {
  if (!aStart || !bStart) return false;
  var aLast = aEnd || aStart;
  var bLast = bEnd || bStart;
  return aStart <= bLast && bStart <= aLast;
}

function core_domainsV2FormatDuration_(days) {
  if (!days) return '';
  var years = Math.floor(days / 365.25);
  var months = Math.floor((days - (years * 365.25)) / 30.44);
  if (years && months) return years + ' anos, ' + months + ' meses';
  if (years) return years + ' anos';
  if (months) return months + ' meses';
  return days + ' dias';
}

function core_domainsV2EffectiveMemberIntervals_(vinculos, today) {
  return (vinculos || []).filter(function(vinculo) {
    return core_domainsV2NormalizeTipoVinculo_(vinculo.TIPO_VINCULO) === 'MEMBRO_EFETIVO';
  }).map(function(vinculo) {
    var start = core_domainsV2Date_(vinculo.DATA_INICIO);
    var end = core_domainsV2Date_(vinculo.DATA_FIM);
    if (!end && core_domainsV2Active_(vinculo)) end = today;
    return start ? { start: start, end: end || start } : null;
  }).filter(function(interval) {
    return !!interval;
  });
}

function core_domainsV2CountIntervalDays_(intervals) {
  return (intervals || []).reduce(function(total, interval) {
    return total + core_domainsV2DaysBetween_(interval.start, interval.end);
  }, 0);
}

function core_domainsV2CountSemestersForIntervals_(vigenciasData, intervals) {
  if (!intervals.length) return '';
  var records = ((vigenciasData.SEMESTRES && vigenciasData.SEMESTRES.records) || []);
  if (!records.length) records = ((vigenciasData.PERIODOS && vigenciasData.PERIODOS.records) || []);
  var seen = {};
  records.forEach(function(record) {
    var start = core_domainsV2Date_(record.DATA_INICIO);
    var end = core_domainsV2Date_(record.DATA_FIM);
    if (!start || !end) return;
    var crosses = intervals.some(function(interval) {
      return core_domainsV2IntervalsOverlap_(interval.start, interval.end, start, end);
    });
    if (crosses) {
      var key = record.ID_SEMESTRE || record.ID_PERIODO || record.NOME_PERIODO || [record.ANO, record.SEMESTRE].join('/');
      seen[String(key || record.__rowNumber || '').trim()] = true;
    }
  });
  return Object.keys(seen).length || '';
}

function core_domainsV2GetRga_(pessoasData, idPessoa, detalhes) {
  var fallback = detalhes && detalhes.RGA ? detalhes.RGA : '';
  var records = (pessoasData.PESSOAS_IDENTIFICADORES && pessoasData.PESSOAS_IDENTIFICADORES.records) || [];
  for (var i = records.length - 1; i >= 0; i--) {
    var record = records[i];
    if (String(record.ID_PESSOA || '').trim() !== idPessoa) continue;
    if (core_domainsV2AuditStatus_(record.TIPO_IDENTIFICADOR) !== 'RGA') continue;
    if (!core_domainsV2AuditIsSim_(record.ATIVO)) continue;
    if (core_domainsV2AuditIsSim_(record.PRINCIPAL)) return record.VALOR_IDENTIFICADOR || fallback;
    fallback = fallback || record.VALOR_IDENTIFICADOR;
  }
  return fallback || '';
}

function core_domainsV2CurrentSemesterFromVigencias_(vigenciasData, refDate) {
  var now = refDate || new Date();
  var semestres = (vigenciasData.SEMESTRES && vigenciasData.SEMESTRES.records) || [];
  var next = null;
  for (var i = 0; i < semestres.length; i++) {
    var record = semestres[i];
    var start = core_domainsV2Date_(record.DATA_INICIO);
    var end = core_domainsV2Date_(record.DATA_FIM);
    var id = String(record.ID_SEMESTRE || '').trim();
    if (!id || !start) continue;
    if (core_domainsV2IntervalsOverlap_(now, now, start, end || start)) {
      return { id: id, startDate: start, endDate: end };
    }
    if (start > now && (!next || start < next.startDate)) {
      next = { id: id, startDate: start, endDate: end };
    }
  }
  return next;
}

function core_domainsV2StudentCurrentSemesterFromRga_(rga, vigenciasData, refDate) {
  var entry = core_parseEntrySemesterFromRga_(rga);
  if (!entry) return null;
  var currentSemester = core_domainsV2CurrentSemesterFromVigencias_(vigenciasData, refDate);
  if (!currentSemester || !currentSemester.id) return null;
  var current = core_parseSemesterId_(currentSemester.id);
  if (!current) return null;
  var diff = (current.year - entry.year) * 2 + (current.semester - entry.semester);
  var semesterNumber = diff + 1;
  return semesterNumber < 1 ? null : semesterNumber;
}

function core_domainsV2BuildMembrosDetalhesSemesterRows_(pessoasData, vigenciasData, opts, report) {
  var refDate = opts.refDate || new Date();
  var detalhes = (pessoasData.MEMBROS_DETALHES && pessoasData.MEMBROS_DETALHES.records) || [];
  var out = [];
  detalhes.forEach(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    if (opts.idPessoa && idPessoa !== opts.idPessoa) return;
    var rga = core_domainsV2GetRga_(pessoasData, idPessoa, record);
    var semester = core_domainsV2StudentCurrentSemesterFromRga_(rga, vigenciasData, refDate);
    if (!rga) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'MEMBRO_DETALHE_SEM_RGA', 'MEMBROS_DETALHES sem RGA para calcular SEMESTRE_ATUAL.', { idPessoa: idPessoa });
    } else if (semester === null) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'RGA_SEM_SEMESTRE_CALCULAVEL', 'RGA nao permitiu calcular SEMESTRE_ATUAL.', { idPessoa: idPessoa, rga: rga });
    }
    out.push({
      rowNumber: record.__rowNumber,
      idPessoa: idPessoa,
      rga: rga,
      semestreAtualAnterior: record.SEMESTRE_ATUAL || '',
      semestreAtualCalculado: semester === null ? '' : semester
    });
  });
  return out;
}

function coreRecalcularMembrosDetalhesSemestreAtualV2_(options) {
  options = options || {};
  var opts = core_domainsV2ResumoOptions_(options);
  opts.idPessoa = options.idPessoa ? String(options.idPessoa).trim() : '';
  opts.refDate = options.refDate ? core_domainsV2Date_(options.refDate) : new Date();
  if (!opts.dryRun && opts.confirmacao !== 'RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2') {
    throw new Error('Para escrever MEMBROS_DETALHES.SEMESTRE_ATUAL, informe confirmacao: "RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2".');
  }

  var report = core_domainsV2AuditNewReport_('RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2');
  report.dryRun = opts.dryRun;
  report.options = {
    idPessoa: opts.idPessoa || null,
    refDate: opts.refDate
  };
  var pessoasData = core_domainsV2OpenPessoas_(report);
  var vigenciasData = core_domainsV2OpenVigencias_(report);
  if (report.totalErros) return report;

  var rows = core_domainsV2BuildMembrosDetalhesSemesterRows_(pessoasData, vigenciasData, opts, report);
  var calculaveis = rows.filter(function(row) {
    return row.semestreAtualCalculado !== '';
  });
  var alteradas = calculaveis.filter(function(row) {
    return String(row.semestreAtualAnterior || '').trim() !== String(row.semestreAtualCalculado);
  });
  report.resumoQuantitativo = {
    linhasAnalisadas: rows.length,
    linhasCalculaveis: calculaveis.length,
    linhasComMudanca: alteradas.length,
    linhasAtualizadas: opts.dryRun ? 0 : alteradas.length
  };
  report.amostra = rows.slice(0, 10);

  if (opts.dryRun) {
    core_domainsV2AuditRecommendation_(report, 'Conferir a amostra; se estiver correta, rodar coreRecalcularMembrosDetalhesSemestreAtualV2({ dryRun: false, confirmacao: "RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2" }).');
    return report;
  }

  var headers = pessoasData.MEMBROS_DETALHES.headers || [];
  var headerMap = core_buildHeaderIndexMap_(headers);
  var semesterCol = core_findHeaderIndex_(headerMap, 'SEMESTRE_ATUAL');
  if (semesterCol < 0) {
    throw new Error('Cabecalho SEMESTRE_ATUAL nao encontrado em MEMBROS_DETALHES.');
  }
  var sheet = pessoasData.MEMBROS_DETALHES.sheet;
  alteradas.forEach(function(row) {
    if (!row.rowNumber) return;
    sheet.getRange(row.rowNumber, semesterCol + 1).setValue(row.semestreAtualCalculado);
  });
  return report;
}

function core_domainsV2ActivePortalExceptions_(pessoasData, idPessoa, today) {
  return ((pessoasData.PORTAL_ACESSOS_EXCECOES && pessoasData.PORTAL_ACESSOS_EXCECOES.records) || []).filter(function(record) {
    if (String(record.ID_PESSOA || '').trim() !== idPessoa) return false;
    var status = core_domainsV2AuditStatus_(record.STATUS);
    if (['REVOGADO', 'ENCERRADO', 'INATIVO', 'EXPIRADO'].indexOf(status) >= 0) return false;
    var start = core_domainsV2Date_(record.DATA_INICIO);
    var end = core_domainsV2Date_(record.DATA_FIM);
    if (start && start > today) return false;
    if (end && end < today) return false;
    return true;
  });
}

function core_domainsV2BuildPortalState_(vinculo, vigResumo, exceptions) {
  var tipo = core_domainsV2NormalizeTipoVinculo_(vinculo.TIPO_VINCULO);
  var active = core_domainsV2Active_(vinculo);
  var profile = String(vigResumo.PERFIS_PORTAL_CALCULADOS || '').trim();
  var allow = false;
  var block = false;
  (exceptions || []).forEach(function(record) {
    var status = core_domainsV2AuditStatus_(record.STATUS);
    if (['BLOQUEADO', 'NEGADO', 'SUSPENSO'].indexOf(status) >= 0) block = true;
    if (String(record.PERFIL_EXTRA || '').trim()) {
      profile = profile ? profile + '; ' + record.PERFIL_EXTRA : record.PERFIL_EXTRA;
      allow = true;
    }
    if (String(record.PERMISSAO_EXTRA || '').trim()) allow = true;
  });
  if (!profile && tipo === 'MEMBRO_EFETIVO' && active) profile = 'MEMBRO';
  if (!profile && allow) profile = 'EXCECAO_PORTAL';
  if (block) return { perfil: profile, portalAtivo: 'NAO' };
  if (allow && profile) return { perfil: profile, portalAtivo: 'SIM' };
  if (tipo === 'MEMBRO_EFETIVO' && active && profile) return { perfil: profile, portalAtivo: 'SIM' };
  return { perfil: profile, portalAtivo: 'NAO' };
}

function core_domainsV2ReadRecordsByKeySoft_(key, unavailable) {
  try {
    return core_readRecordsByKey_(key, { skipBlankRows: true }) || [];
  } catch (err) {
    unavailable[key] = err && err.message ? err.message : String(err);
    return [];
  }
}

function core_domainsV2ActivityData_() {
  var unavailable = {};
  return {
    apresentacoes: core_domainsV2ReadRecordsByKeySoft_('ATIVIDADES_V2_APRESENTACOES', unavailable),
    portalAtividadesDetalhes: core_domainsV2ReadRecordsByKeySoft_('ATIVIDADES_V2_PORTAL_ATIVIDADES_DETALHES', unavailable),
    presencas: core_domainsV2ReadRecordsByKeySoft_('ATIVIDADES_V2_PRESENCAS_REGISTROS', unavailable),
    portalFrequencia: core_domainsV2ReadRecordsByKeySoft_('ATIVIDADES_V2_PORTAL_FREQUENCIA_MEMBROS', unavailable),
    unavailable: unavailable
  };
}

function core_domainsV2ParseJsonArray_(value) {
  if (Array.isArray(value)) return value;
  var text = String(value || '').trim();
  if (!text) return [];
  try {
    var parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function core_domainsV2PresentationRecordsFromPortalDetails_(details) {
  var records = [];
  (details || []).forEach(function(detail) {
    var presentations = core_domainsV2ParseJsonArray_(core_domainsV2LegacyValue_(detail, [
      'APRESENTACOES_PUBLICAS_JSON',
      'apresentacoesPublicas'
    ]));
    if (!presentations.length) return;

    var base = {
      ID_ATIVIDADE: core_domainsV2LegacyValue_(detail, ['ID_ATIVIDADE', 'idAtividade']),
      DATA_ATIVIDADE: core_domainsV2LegacyValue_(detail, ['DATA_ATIVIDADE', 'dataAtividade', 'DATA']),
      PERIODO: core_domainsV2LegacyValue_(detail, ['PERIODO', 'ID_PERIODO', 'periodo'])
    };
    presentations.forEach(function(item) {
      if (!item || typeof item !== 'object' || Array.isArray(item)) return;
      var out = {};
      Object.keys(item).forEach(function(key) {
        out[key] = item[key];
      });
      Object.keys(base).forEach(function(key) {
        if (out[key] === undefined || out[key] === null || out[key] === '') out[key] = base[key];
      });
      records.push(out);
    });
  });
  return records;
}

function core_domainsV2RecordMatchesPessoa_(record, ctx) {
  var id = String(core_domainsV2LegacyValue_(record, [
    'ID_PESSOA',
    'ID_PESSOA_APRESENTADOR',
    'ID_PESSOA_PRINCIPAL',
    'idPessoa',
    'idPessoaApresentador',
    'idPessoaPrincipal',
    'Id Pessoa'
  ]) || '').trim();
  var rga = core_domainsV2Rga_(core_domainsV2LegacyValue_(record, ['RGA', 'rga', 'rgaApresentador']));
  var email = core_domainsV2Email_(core_domainsV2LegacyValue_(record, [
    'EMAIL',
    'EMAIL_APRESENTADOR',
    'Email',
    'E-mail',
    'email',
    'emailApresentador'
  ]));
  return (!!ctx.idPessoa && id === ctx.idPessoa) || (!!ctx.rga && rga === ctx.rga) || (!!ctx.email && email === ctx.email);
}

function core_domainsV2StatusConcluido_(value) {
  var status = core_domainsV2AuditStatus_(value);
  return ['REALIZADA', 'CONCLUIDA', 'CONCLUÍDA', 'FINALIZADA', 'APROVADA', 'CONFIRMADA'].indexOf(status) >= 0;
}

function core_domainsV2StatusPendente_(value) {
  var status = core_domainsV2AuditStatus_(value);
  return ['PENDENTE', 'AGENDADA', 'EM_ABERTO', 'ABERTA', 'SOLICITADA', 'PLANEJADA'].indexOf(status) >= 0;
}

function core_domainsV2PresentationSummary_(activityData, ctx) {
  var records = activityData.apresentacoes && activityData.apresentacoes.length
    ? activityData.apresentacoes
    : core_domainsV2PresentationRecordsFromPortalDetails_(activityData.portalAtividadesDetalhes || []);
  var concluded = [];
  var pending = false;
  var filePending = false;
  records.forEach(function(record) {
    if (!core_domainsV2RecordMatchesPessoa_(record, ctx)) return;
    var status = core_domainsV2LegacyValue_(record, ['STATUS_APRESENTACAO', 'STATUS', 'SITUACAO']);
    var fileStatus = core_domainsV2LegacyValue_(record, ['STATUS_ARQUIVO', 'ARQUIVO_STATUS', 'STATUS_DRIVE']);
    if (core_domainsV2StatusConcluido_(status)) concluded.push(record);
    if (core_domainsV2StatusPendente_(status)) pending = true;
    if (core_domainsV2StatusPendente_(fileStatus)) filePending = true;
  });
  concluded.sort(function(a, b) {
    var da = core_domainsV2Date_(core_domainsV2LegacyValue_(a, ['DATA_ATIVIDADE', 'DATA_APRESENTACAO', 'DATA_REALIZADA', 'DATA', 'DATA_INICIO', 'dataAtividade', 'dataApresentacao']));
    var db = core_domainsV2Date_(core_domainsV2LegacyValue_(b, ['DATA_ATIVIDADE', 'DATA_APRESENTACAO', 'DATA_REALIZADA', 'DATA', 'DATA_INICIO', 'dataAtividade', 'dataApresentacao']));
    return core_domainsV2ResumoTimestamp_(db) - core_domainsV2ResumoTimestamp_(da);
  });
  var last = concluded[0] || {};
  return {
    count: concluded.length,
    lastPeriod: core_domainsV2LegacyValue_(last, ['PERIODO', 'CICLO', 'SEMESTRE', 'ID_PERIODO', 'PERIODO_REFERENCIA']),
    hasPending: pending,
    hasFilePending: filePending
  };
}

function core_domainsV2FrequencySummary_(activityData, ctx, periodoReferencia) {
  var portal = activityData.portalFrequencia || [];
  for (var i = 0; i < portal.length; i++) {
    var record = portal[i];
    if (!core_domainsV2RecordMatchesPessoa_(record, ctx)) continue;
    var periodo = core_domainsV2LegacyValue_(record, ['PERIODO', 'PERIODO_REFERENCIA', 'ID_PERIODO', 'SEMESTRE']);
    if (periodoReferencia && String(periodo || '').trim() !== String(periodoReferencia || '').trim()) continue;
    var ready = core_domainsV2LegacyValue_(record, ['FREQUENCIA_RESUMIDA', 'RESUMO_FREQUENCIA', 'FREQUENCIA', 'STATUS_FREQUENCIA']);
    if (ready) return String(ready);
    var percentual = core_domainsV2LegacyValue_(record, ['PERCENTUAL_FREQUENCIA', 'FREQUENCIA_PERCENTUAL']);
    var presencas = core_domainsV2LegacyValue_(record, ['PRESENCAS', 'QTD_PRESENCAS']);
    var faltas = core_domainsV2LegacyValue_(record, ['FALTAS', 'QTD_FALTAS']);
    var justificadas = core_domainsV2LegacyValue_(record, ['JUSTIFICADAS', 'FALTAS_JUSTIFICADAS']);
    return [
      percentual ? 'Frequencia ' + percentual : '',
      presencas !== '' ? 'Presencas ' + presencas : '',
      faltas !== '' ? 'Faltas ' + faltas : '',
      justificadas !== '' ? 'Justificadas ' + justificadas : ''
    ].filter(String).join('; ');
  }
  var counts = { presenca: 0, falta: 0, justificada: 0 };
  (activityData.presencas || []).forEach(function(record) {
    if (!core_domainsV2RecordMatchesPessoa_(record, ctx)) return;
    var periodo = core_domainsV2LegacyValue_(record, ['PERIODO', 'PERIODO_REFERENCIA', 'ID_PERIODO', 'SEMESTRE']);
    if (periodoReferencia && String(periodo || '').trim() !== String(periodoReferencia || '').trim()) return;
    var status = core_domainsV2AuditStatus_(core_domainsV2LegacyValue_(record, ['STATUS_PRESENCA', 'STATUS', 'SITUACAO']));
    if (status.indexOf('JUST') >= 0) counts.justificada++;
    else if (status.indexOf('FALTA') >= 0 || status === 'AUSENTE') counts.falta++;
    else if (status.indexOf('PRES') >= 0 || status === 'PRESENTE') counts.presenca++;
  });
  if (!counts.presenca && !counts.falta && !counts.justificada) return '';
  return 'Presencas ' + counts.presenca + '; Faltas ' + counts.falta + '; Justificadas ' + counts.justificada;
}

function core_domainsV2HasCriticalFrequency_(frequencyText) {
  var text = core_domainsV2AuditStatus_(frequencyText);
  return text.indexOf('CRITICA') >= 0 || text.indexOf('CRÍTICA') >= 0 || text.indexOf('BAIXA') >= 0;
}

function core_domainsV2SuspensionFlag_(eventos, idPessoa) {
  return (eventos || []).some(function(record) {
    if (String(record.ID_PESSOA || '').trim() !== idPessoa) return false;
    var tipo = core_domainsV2AuditStatus_(record.TIPO_EVENTO);
    var status = core_domainsV2AuditStatus_(record.STATUS_EVENTO);
    return tipo.indexOf('SUSPEN') >= 0 && ['HOMOLOGADO', 'PROCESSADO', 'CONCLUIDO', 'CONCLUÍDO', 'APROVADO'].indexOf(status) >= 0;
  }) ? 'SIM' : 'NAO';
}

function core_domainsV2Eligibility_(vinculo) {
  var tipo = core_domainsV2NormalizeTipoVinculo_(vinculo.TIPO_VINCULO);
  var active = core_domainsV2Active_(vinculo);
  if (tipo === 'MEMBRO_EFETIVO' && active) return 'ELEGIVEL';
  if (tipo === 'MEMBRO_EM_ESPERA') return 'INELEGIVEL_MEMBRO_EM_ESPERA';
  if (tipo === 'EGRESSO') return 'INELEGIVEL_EX_MEMBRO';
  if (!active) return 'INELEGIVEL_SEM_VINCULO_ATIVO';
  return 'PENDENTE_ANALISE';
}

function core_domainsV2BuildPessoasResumoRows_(pessoasData, vigenciasData, options, report) {
  options = options || {};
  var today = new Date();
  var detalhesById = core_domainsV2IndexFirstBy_((pessoasData.MEMBROS_DETALHES && pessoasData.MEMBROS_DETALHES.records) || [], 'ID_PESSOA');
  var vinculosById = core_domainsV2IndexManyBy_((pessoasData.VINCULOS_GEAPA && pessoasData.VINCULOS_GEAPA.records) || [], 'ID_PESSOA');
  var vigResumoById = core_domainsV2IndexFirstBy_((vigenciasData.VIGENCIAS_RESUMO_ATUAL && vigenciasData.VIGENCIAS_RESUMO_ATUAL.records) || [], 'ID_PESSOA');
  var eventos = (pessoasData.MEMBROS_EVENTOS_VINCULO && pessoasData.MEMBROS_EVENTOS_VINCULO.records) || [];
  var activityData = core_domainsV2ActivityData_();
  var unavailable = activityData.unavailable || {};
  unavailable.DATA_LIMITE_ESTIMADA_DIRETORIA = 'Regra de limite de diretoria ainda nao esta segura para calculo automatico.';

  var rows = ((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || []).filter(function(pessoa) {
    return !options.idPessoa || String(pessoa.ID_PESSOA || '').trim() === String(options.idPessoa || '').trim();
  }).map(function(pessoa) {
    var idPessoa = String(pessoa.ID_PESSOA || '').trim();
    var detalhes = detalhesById[idPessoa] || {};
    var vinculos = vinculosById[idPessoa] || [];
    var vinculo = core_domainsV2PickCurrentVinculo_(vinculos) || {};
    var vigResumo = vigResumoById[idPessoa] || {};
    var rga = core_domainsV2GetRga_(pessoasData, idPessoa, detalhes);
    var email = pessoa.EMAIL_PRINCIPAL || '';
    var ctx = { idPessoa: idPessoa, rga: core_domainsV2Rga_(rga), email: core_domainsV2Email_(email) };
    var intervals = core_domainsV2EffectiveMemberIntervals_(vinculos, today);
    var presentation = core_domainsV2PresentationSummary_(activityData, ctx);
    var frequency = core_domainsV2FrequencySummary_(activityData, ctx, options.periodoReferencia);
    var portal = core_domainsV2BuildPortalState_(vinculo, vigResumo, core_domainsV2ActivePortalExceptions_(pessoasData, idPessoa, today));
    var tipoAtual = vinculo.TIPO_VINCULO || '';
    var normalizedTipoAtual = core_domainsV2NormalizeTipoVinculo_(tipoAtual);
    var activeCurrent = core_domainsV2Active_(vinculo);
    var cargoAtual = vigResumo.CARGO_FUNCAO_ATUAL || '';
    var pendencias = [];
    if (!email) pendencias.push('SEM_EMAIL');
    if (normalizedTipoAtual === 'MEMBRO_EFETIVO' && !rga) pendencias.push('SEM_RGA');
    if (!pessoa.NOME_EXIBICAO && !pessoa.NOME_COMPLETO) pendencias.push('CADASTRO_INCOMPLETO');
    if (presentation.hasPending) pendencias.push('APRESENTACAO_PENDENTE');
    if (presentation.hasFilePending) pendencias.push('ARQUIVO_APRESENTACAO_PENDENTE');
    if (core_domainsV2HasCriticalFrequency_(frequency)) pendencias.push('FREQUENCIA_CRITICA');
    if (frequency && core_domainsV2AuditStatus_(frequency).indexOf('JUSTIFICATIVA_PENDENTE') >= 0) pendencias.push('JUSTIFICATIVA_PENDENTE');
    if (!cargoAtual && normalizedTipoAtual === 'MEMBRO_EFETIVO' && activeCurrent) cargoAtual = 'Membro';

    return {
      ID_PESSOA: idPessoa,
      RGA: rga,
      NOME_EXIBICAO: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || '',
      EMAIL: email,
      TIPO_VINCULO_ATUAL: tipoAtual,
      STATUS_VINCULO_ATUAL: vinculo.STATUS_VINCULO || '',
      CARGO_FUNCAO_ATUAL: cargoAtual,
      PERFIL_PORTAL_CALCULADO: portal.perfil,
      PORTAL_ATIVO: portal.portalAtivo,
      TEMPO_EFETIVO_NO_GRUPO: core_domainsV2FormatDuration_(core_domainsV2CountIntervalDays_(intervals)),
      QTD_SEMESTRES_NO_GRUPO: core_domainsV2CountSemestersForIntervals_(vigenciasData, intervals),
      QTD_APRESENTACOES_REALIZADAS: presentation.count,
      PERIODO_ULTIMA_APRESENTACAO: presentation.lastPeriod || '',
      FREQUENCIA_RESUMIDA: frequency || '',
      PENDENCIAS_ABERTAS: pendencias.length ? pendencias.join('; ') : 'SEM_PENDENCIAS',
      FLAG_JA_FOI_SUSPENSO: core_domainsV2SuspensionFlag_(eventos, idPessoa),
      STATUS_ELEGIBILIDADE_DIRETORIA: core_domainsV2Eligibility_(vinculo),
      DATA_LIMITE_ESTIMADA_DIRETORIA: '',
      ULTIMA_ATUALIZACAO: new Date()
    };
  });
  report.camposNaoCalculaveis = Object.keys(unavailable).map(function(key) {
    return { campo: key, motivo: unavailable[key] };
  });
  return rows;
}

function coreRecalcularPessoasResumoOperacionalV2_(options) {
  var opts = core_domainsV2ResumoOptions_(options);
  opts.idPessoa = options && options.idPessoa ? String(options.idPessoa).trim() : '';
  opts.periodoReferencia = options && options.periodoReferencia ? String(options.periodoReferencia).trim() : '';
  if (!opts.dryRun && opts.confirmacao !== 'RECALCULAR_PESSOAS_RESUMO_V2') {
    throw new Error('Para escrever PESSOAS_RESUMO_OPERACIONAL, informe confirmacao: "RECALCULAR_PESSOAS_RESUMO_V2".');
  }

  var report = core_domainsV2AuditNewReport_('RECALCULAR_PESSOAS_RESUMO_OPERACIONAL_V2');
  report.dryRun = opts.dryRun;
  var pessoasData = core_domainsV2OpenPessoas_(report);
  var vigenciasData = core_domainsV2OpenVigencias_(report);
  if (report.totalErros) return report;

  var rows = core_domainsV2BuildPessoasResumoRows_(pessoasData, vigenciasData, opts, report);
  var headers = pessoasData.PESSOAS_RESUMO_OPERACIONAL.headers || [];
  var stats = core_domainsV2PessoasResumoStats_(rows);
  report.resumoQuantitativo = {
    totalPessoasAnalisadas: rows.length,
    totalMembrosAtivos: stats.totalMembrosAtivos,
    totalExMembros: stats.totalExMembros,
    totalMembrosEmEspera: stats.totalMembrosEmEspera,
    pessoasBase: (pessoasData.PESSOAS_BASE.records || []).length,
    linhasCalculadas: rows.length,
    linhasExistentesAntes: (pessoasData.PESSOAS_RESUMO_OPERACIONAL.records || []).length,
    resumosAtualizados: opts.dryRun ? 0 : rows.length
  };
  report.options = {
    idPessoa: opts.idPessoa || null,
    periodoReferencia: opts.periodoReferencia || null
  };
  core_domainsV2AttachLegacyComparisonSummary_(report);

  if (opts.dryRun) {
    report.amostra = rows.slice(0, 10);
    core_domainsV2AuditRecommendation_(report, 'Conferir a amostra; se estiver correta, rodar coreRecalcularPessoasResumoOperacionalV2({ dryRun: false, confirmacao: "RECALCULAR_PESSOAS_RESUMO_V2" }).');
    return report;
  }

  if (rows.length) {
    core_domainsV2WritePessoasResumoRows_(pessoasData.PESSOAS_RESUMO_OPERACIONAL, headers, rows, opts);
  }
  report.resumoQuantitativo.linhasEscritas = rows.length;
  return report;
}

function core_domainsV2AttachLegacyComparisonSummary_(report) {
  try {
    var comparison = coreCompararLegadoComV2_({ limit: 10 });
    report.divergenciasLegado = {
      totalErros: comparison.totalErros,
      totalAvisos: comparison.totalAvisos,
      resumoQuantitativo: comparison.resumoQuantitativo || {},
      avisos: (comparison.avisos || []).slice(0, 10)
    };
  } catch (err) {
    report.divergenciasLegado = {
      comparavel: false,
      motivo: err && err.message ? err.message : String(err)
    };
  }
}

function core_domainsV2PessoasResumoStats_(rows) {
  var stats = { totalMembrosAtivos: 0, totalExMembros: 0, totalMembrosEmEspera: 0 };
  (rows || []).forEach(function(row) {
    var tipo = core_domainsV2NormalizeTipoVinculo_(row.TIPO_VINCULO_ATUAL);
    var status = core_domainsV2AuditStatus_(row.STATUS_VINCULO_ATUAL);
    if (tipo === 'MEMBRO_EFETIVO' && (status === 'ATIVO' || row.PORTAL_ATIVO === 'SIM')) stats.totalMembrosAtivos++;
    if (tipo === 'EGRESSO') stats.totalExMembros++;
    if (tipo === 'MEMBRO_EM_ESPERA') stats.totalMembrosEmEspera++;
  });
  return stats;
}

function core_domainsV2WritePessoasResumoRows_(targetData, headers, rows, opts) {
  var sheet = targetData.sheet;
  if (opts.idPessoa) {
    var existingById = core_domainsV2IndexFirstBy_(targetData.records || [], 'ID_PESSOA');
    rows.forEach(function(row) {
      var existingRow = existingById[String(row.ID_PESSOA || '').trim()];
      if (existingRow && existingRow.__rowNumber) {
        sheet.getRange(existingRow.__rowNumber, 1, 1, headers.length).setValues([core_buildRowFromObjectByHeaders_(headers, row)]);
      } else {
        sheet.appendRow(core_buildRowFromObjectByHeaders_(headers, row));
      }
    });
    return rows.length;
  }

  var existing = (targetData.records || []).map(function(record) {
    return core_domainsV2CloneRecord_(record);
  });
  var existingIndexById = {};
  existing.forEach(function(record, index) {
    var id = String(record.ID_PESSOA || '').trim();
    if (id) existingIndexById[id] = index;
  });
  var appendRows = [];
  rows.forEach(function(row) {
    var id = String(row.ID_PESSOA || '').trim();
    if (Object.prototype.hasOwnProperty.call(existingIndexById, id)) {
      existing[existingIndexById[id]] = row;
    } else {
      appendRows.push(row);
    }
  });
  if (existing.length) {
    sheet.getRange(2, 1, existing.length, headers.length).setValues(existing.map(function(row) {
      return core_buildRowFromObjectByHeaders_(headers, row);
    }));
  }
  appendRows.forEach(function(row) {
    sheet.appendRow(core_buildRowFromObjectByHeaders_(headers, row));
  });
  return rows.length;
}

function coreDiagnosticarPessoasResumoOperacionalV2_(options) {
  options = options || {};
  var report = core_domainsV2AuditNewReport_('DIAGNOSTICAR_PESSOAS_RESUMO_OPERACIONAL_V2');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  var vigenciasData = core_domainsV2OpenVigencias_(report);
  if (report.totalErros) return report;

  var baseById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [], 'ID_PESSOA');
  var detalhesById = core_domainsV2IndexFirstBy_((pessoasData.MEMBROS_DETALHES && pessoasData.MEMBROS_DETALHES.records) || [], 'ID_PESSOA');
  var vinculosById = core_domainsV2IndexManyBy_((pessoasData.VINCULOS_GEAPA && pessoasData.VINCULOS_GEAPA.records) || [], 'ID_PESSOA');
  var resumoById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [], 'ID_PESSOA');
  var vigResumoById = core_domainsV2IndexFirstBy_((vigenciasData.VIGENCIAS_RESUMO_ATUAL && vigenciasData.VIGENCIAS_RESUMO_ATUAL.records) || [], 'ID_PESSOA');

  Object.keys(baseById).forEach(function(idPessoa) {
    if (!resumoById[idPessoa]) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'PESSOA_SEM_RESUMO', 'Pessoa sem linha em PESSOAS_RESUMO_OPERACIONAL.', { idPessoa: idPessoa });
    }
  });

  Object.keys(resumoById).forEach(function(idPessoa) {
    var resumo = resumoById[idPessoa];
    var pessoa = baseById[idPessoa];
    var detalhes = detalhesById[idPessoa] || {};
    var vinculo = core_domainsV2PickCurrentVinculo_(vinculosById[idPessoa] || []) || {};
    var tipo = core_domainsV2NormalizeTipoVinculo_(vinculo.TIPO_VINCULO || resumo.TIPO_VINCULO_ATUAL);
    var active = core_domainsV2Active_(vinculo);
    var vigResumo = vigResumoById[idPessoa] || {};
    if (!pessoa) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'RESUMO_SEM_PESSOA', 'Resumo operacional aponta para pessoa inexistente.', { idPessoa: idPessoa });
      return;
    }
    if (tipo === 'MEMBRO_EFETIVO' && active && resumo.PORTAL_ATIVO !== 'SIM') {
      core_domainsV2AuditIssue_(report, 'AVISO', 'MEMBRO_ATIVO_SEM_PORTAL_ATIVO', 'Membro efetivo ativo sem PORTAL_ATIVO=SIM.', { idPessoa: idPessoa });
    }
    if (tipo === 'EGRESSO' && resumo.PORTAL_ATIVO === 'SIM') {
      core_domainsV2AuditIssue_(report, 'AVISO', 'EGRESSO_COM_PORTAL_ATIVO', 'Egresso com PORTAL_ATIVO=SIM; verificar excecao formal.', { idPessoa: idPessoa });
    }
    if (tipo === 'MEMBRO_EFETIVO' && active && !core_domainsV2GetRga_(pessoasData, idPessoa, detalhes)) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'MEMBRO_ATIVO_SEM_RGA', 'Membro efetivo ativo sem RGA.', { idPessoa: idPessoa });
    }
    if (tipo === 'MEMBRO_EFETIVO' && active && !pessoa.EMAIL_PRINCIPAL) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'MEMBRO_ATIVO_SEM_EMAIL', 'Membro efetivo ativo sem e-mail principal.', { idPessoa: idPessoa });
    }
    if (vigResumo.CARGO_FUNCAO_ATUAL && resumo.CARGO_FUNCAO_ATUAL !== vigResumo.CARGO_FUNCAO_ATUAL) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'CARGO_ATUAL_DIVERGENTE_VIGENCIAS', 'Cargo atual no resumo de Pessoas diverge de VIGENCIAS_RESUMO_ATUAL.', {
        idPessoa: idPessoa,
        pessoas: resumo.CARGO_FUNCAO_ATUAL,
        vigencias: vigResumo.CARGO_FUNCAO_ATUAL
      });
    }
    if (vigResumo.PERFIS_PORTAL_CALCULADOS && String(resumo.PERFIL_PORTAL_CALCULADO || '').indexOf(String(vigResumo.PERFIS_PORTAL_CALCULADOS || '')) < 0) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'PERFIL_PORTAL_DIVERGENTE_VIGENCIAS', 'Perfil portal no resumo de Pessoas diverge de VIGENCIAS_RESUMO_ATUAL.', {
        idPessoa: idPessoa,
        pessoas: resumo.PERFIL_PORTAL_CALCULADO,
        vigencias: vigResumo.PERFIS_PORTAL_CALCULADOS
      });
    }
    ['QTD_SEMESTRES_NO_GRUPO', 'QTD_APRESENTACOES_REALIZADAS', 'FREQUENCIA_RESUMIDA', 'PENDENCIAS_ABERTAS'].forEach(function(field) {
      if (String(resumo[field] || '').trim() === '') {
        core_domainsV2AuditIssue_(report, 'AVISO', field + '_VAZIO', field + ' vazio em PESSOAS_RESUMO_OPERACIONAL.', { idPessoa: idPessoa });
      }
    });
  });

  report.resumoQuantitativo = {
    pessoasBase: Object.keys(baseById).length,
    resumos: Object.keys(resumoById).length,
    totalErros: report.totalErros,
    totalAvisos: report.totalAvisos
  };
  return report;
}

function core_domainsV2ReadLegacyRecordsSafe_(key) {
  try {
    return core_readRecordsByKey_(key, { skipBlankRows: true }) || [];
  } catch (err) {
    return [];
  }
}

function core_domainsV2LegacyValue_(record, headers) {
  var keys = Object.keys(record || {});
  for (var i = 0; i < headers.length; i++) {
    var target = core_normalizeHeader_(headers[i]);
    for (var j = 0; j < keys.length; j++) {
      if (core_normalizeHeader_(keys[j]) === target) return record[keys[j]];
    }
  }
  return '';
}

function core_domainsV2CompareLegacyList_(report, label, legacyRecords, pessoasData, expectedType, limit) {
  var missing = [];
  legacyRecords.forEach(function(record) {
    var email = core_domainsV2Email_(core_domainsV2LegacyValue_(record, ['EMAIL', 'Email', 'E-mail']));
    var rga = core_domainsV2Rga_(core_domainsV2LegacyValue_(record, ['RGA']));
    var idPessoa = rga
      ? core_domainsV2FindPessoaIdByRga_(pessoasData, rga)
      : core_domainsV2FindPessoaIdByEmail_(pessoasData, email);
    if (!idPessoa) {
      missing.push({ rga: rga, email: email, row: record.__rowNumber });
      return;
    }
    var vinculos = ((pessoasData.VINCULOS_GEAPA && pessoasData.VINCULOS_GEAPA.records) || []).filter(function(vinculo) {
      return String(vinculo.ID_PESSOA || '').trim() === idPessoa &&
        core_domainsV2AuditTipoVinculo_(vinculo.TIPO_VINCULO) === expectedType &&
        core_domainsV2Active_(vinculo);
    });
    if (!vinculos.length) {
      missing.push({ idPessoa: idPessoa, rga: rga, email: email, row: record.__rowNumber, reason: 'SEM_VINCULO_ATIVO_' + expectedType });
    }
  });
  report.resumoQuantitativo[label] = {
    legado: legacyRecords.length,
    divergencias: missing.length
  };
  if (missing.length) {
    core_domainsV2AuditIssue_(report, 'AVISO', 'LEGADO_V2_DIVERGENCIA_' + label, 'Divergencias diagnosticas entre legado e v2.', {
      total: missing.length,
      amostra: missing.slice(0, limit || 20)
    });
  }
}

function coreCompararLegadoComV2_(opts) {
  opts = opts || {};
  var report = core_domainsV2AuditNewReport_('COMPARAR_LEGADO_COM_V2');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) return report;
  var limit = Number(opts.limit || 20);

  core_domainsV2CompareLegacyList_(report, 'MEMBROS_ATUAIS', core_domainsV2ReadLegacyRecordsSafe_('MEMBERS_ATUAIS'), pessoasData, 'MEMBRO_EFETIVO', limit);
  core_domainsV2CompareLegacyList_(report, 'MEMBROS_HIST', core_domainsV2ReadLegacyRecordsSafe_('MEMBERS_HIST'), pessoasData, 'EGRESSO', limit);
  core_domainsV2CompareLegacyList_(report, 'MEMBROS_FUTURO', core_domainsV2ReadLegacyRecordsSafe_('MEMBERS_FUTURO'), pessoasData, 'MEMBRO_EM_ESPERA', limit);
  return report;
}
