/**
 * ============================================================
 * 30_core_domains_v2_audit.js
 * ============================================================
 *
 * Auditorias somente leitura das planilhas centrais v2.
 *
 * Regras:
 * - nao escreve em planilhas;
 * - nao limpa destino v2;
 * - nao altera fontes legadas;
 * - retorna relatorios estruturados para diagnostico pos-migracao.
 */

var CORE_DOMAINS_V2_CONTRACT_KEYS = Object.freeze({
  PESSOAS_V2_BASE: 'PESSOAS_BASE',
  PESSOAS_V2_IDENTIFICADORES: 'PESSOAS_IDENTIFICADORES',
  PESSOAS_V2_MEMBROS_DETALHES: 'MEMBROS_DETALHES',
  PESSOAS_V2_LINKS_PERFIS: 'PESSOAS_LINKS_PERFIS',
  PESSOAS_V2_SOLICITACOES_ATUALIZACAO_CADASTRAL: 'SOLICITACOES_ATUALIZACAO_CADASTRAL',
  PESSOAS_V2_SOLICITACOES_VINCULO: 'SOLICITACOES_VINCULO',
  PESSOAS_V2_COLABORADORES_ACADEMICOS: 'COLABORADORES_ACADEMICOS',
  PESSOAS_V2_PARTICIPANTES_EXTERNOS_DETALHES: 'PARTICIPANTES_EXTERNOS_DETALHES',
  PESSOAS_V2_VINCULOS_GEAPA: 'VINCULOS_GEAPA',
  PESSOAS_V2_MEMBROS_EVENTOS_VINCULO: 'MEMBROS_EVENTOS_VINCULO',
  PESSOAS_V2_COMUNICACAO_CONSENTIMENTOS: 'PESSOAS_COMUNICACAO_CONSENTIMENTOS',
  PESSOAS_V2_PORTAL_ACESSOS_EXCECOES: 'PORTAL_ACESSOS_EXCECOES',
  PESSOAS_V2_RESUMO_OPERACIONAL: 'PESSOAS_RESUMO_OPERACIONAL',
  VIGENCIAS_V2_SEMESTRES: 'SEMESTRES',
  VIGENCIAS_V2_CICLOS: 'CICLOS',
  VIGENCIAS_V2_DIRETORIAS: 'DIRETORIAS',
  VIGENCIAS_V2_SEMESTRES_DIRETORIA: 'SEMESTRES_DIRETORIA',
  VIGENCIAS_V2_CARGOS_CONFIG: 'CARGOS_CONFIG',
  VIGENCIAS_V2_FUNCOES: 'VIGENCIAS_FUNCOES',
  VIGENCIAS_V2_RESUMO_ATUAL: 'VIGENCIAS_RESUMO_ATUAL'
});

function core_getDomainsV2ContractKeys_() {
  return CORE_DOMAINS_V2_CONTRACT_KEYS;
}

function core_domainsV2AuditNewReport_(name) {
  return {
    ok: true,
    dominio: name,
    totalErros: 0,
    totalAvisos: 0,
    erros: [],
    avisos: [],
    recomendacoes: [],
    resumoQuantitativo: {}
  };
}

function core_domainsV2AuditIssue_(report, severity, code, message, details) {
  var item = {
    code: code,
    message: message
  };
  if (details) item.details = details;

  if (severity === 'ERRO') {
    report.erros.push(item);
    report.totalErros++;
    report.ok = false;
  } else {
    report.avisos.push(item);
    report.totalAvisos++;
  }
}

function core_domainsV2AuditRecommendation_(report, message) {
  if (report.recomendacoes.indexOf(message) < 0) {
    report.recomendacoes.push(message);
  }
}

function core_domainsV2AuditNormalize_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function core_domainsV2AuditEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function core_domainsV2AuditDigits_(value) {
  return String(value || '').replace(/\D+/g, '');
}

function core_domainsV2AuditIsSim_(value) {
  var v = core_domainsV2AuditNormalize_(value);
  return v === 'SIM' || v === 'S' || v === 'TRUE' || v === 'ATIVO' || v === 'ATIVA';
}

function core_domainsV2AuditStatus_(value) {
  return core_domainsV2AuditNormalize_(value);
}

function core_domainsV2AuditTipoVinculo_(value) {
  var tipo = core_domainsV2AuditStatus_(value);
  if (tipo === 'EX_MEMBRO' || tipo === 'EX-MEMBRO' || tipo === 'EX MEMBRO') return 'EGRESSO';
  return tipo;
}

function core_domainsV2AuditVigenciaPareceAtiva_(record) {
  var status = core_domainsV2AuditStatus_(record.STATUS_VIGENCIA);
  if (status === 'ATIVA' || status === 'ATIVO') return true;
  if (status === 'ENCERRADA' || status === 'ENCERRADO' || status === 'INATIVA' || status === 'INATIVO') return false;
  return false;
}

function core_domainsV2AuditOpenDomain_(domainKey, report, options) {
  options = options || {};
  var environment = core_normalizeDomainEnv_(options);
  var spreadsheet;
  try {
    spreadsheet = core_openDomainSpreadsheet_(domainKey, {
      ambiente: environment,
      registryRaw: options.registryRaw,
      openSpreadsheetById: options.openSpreadsheetById
    });
  } catch (err) {
    core_domainsV2AuditIssue_(report, 'ERRO', err.code || 'DOMINIO_DB_INDISPONIVEL', 'Dominio v2 indisponivel no Registry.', {
      domain: domainKey,
      ambiente: environment,
      error: err.message
    });
    return {};
  }
  report.ambiente = environment;
  var out = {};
  var schema = CORE_DOMAINS_V2_SCHEMAS[domainKey] || [];

  schema.forEach(function(definition) {
    var sheet = spreadsheet.getSheetByName(definition.sheetName);
    if (!sheet) {
      core_domainsV2AuditIssue_(report, definition.optional === true ? 'AVISO' : 'ERRO', 'ABA_V2_NAO_ENCONTRADA', definition.optional === true ? 'Aba v2 opcional ainda nao foi preparada.' : 'Aba v2 nao encontrada.', {
        sheetName: definition.sheetName,
        optional: definition.optional === true
      });
      out[definition.sheetName] = { sheet: null, records: [], headers: [] };
      return;
    }

    var headers = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(value) {
          return String(value || '').trim();
        })
      : [];
    var headerMap = core_buildHeaderIndexMap_(headers, { normalize: true, oneBased: false, keepFirst: true });
    var readAliases = definition.readAliases || {};
    var optionalHeaders = (definition.optionalHeaders || []).reduce(function(acc, header) {
      acc[core_normalizeHeader_(header)] = true;
      return acc;
    }, {});
    definition.headers.forEach(function(header) {
      var canonicalKey = core_normalizeHeader_(header);
      if (!Object.prototype.hasOwnProperty.call(headerMap, canonicalKey)) {
        if (optionalHeaders[canonicalKey]) return;
        var aliases = readAliases[header] || [];
        var aliasFound = aliases.some(function(alias) {
          return Object.prototype.hasOwnProperty.call(headerMap, core_normalizeHeader_(alias));
        });
        if (aliasFound) {
          core_domainsV2AuditIssue_(report, 'AVISO', 'CABECALHO_V2_ALIAS_LEGADO', 'Cabecalho v2 lido por alias temporario.', {
            sheetName: definition.sheetName,
            header: header
          });
          return;
        }
        core_domainsV2AuditIssue_(report, 'ERRO', 'CABECALHO_V2_AUSENTE', 'Cabecalho esperado ausente.', {
          sheetName: definition.sheetName,
          header: header
        });
      }
    });

    var records = core_readSheetRecords_(sheet, { skipBlankRows: true });
    records.forEach(function(record) {
      Object.keys(readAliases).forEach(function(canonicalHeader) {
        if (String(record[canonicalHeader] == null ? '' : record[canonicalHeader]).trim()) return;
        var aliases = readAliases[canonicalHeader] || [];
        for (var i = 0; i < aliases.length; i++) {
          if (String(record[aliases[i]] == null ? '' : record[aliases[i]]).trim()) {
            record[canonicalHeader] = record[aliases[i]];
            break;
          }
        }
      });
    });

    out[definition.sheetName] = {
      sheet: sheet,
      records: records,
      headers: headers
    };
  });

  return out;
}

function core_domainsV2AuditIndexBy_(records, field) {
  var out = {};
  (records || []).forEach(function(record) {
    var key = String(record[field] || '').trim();
    if (!key) return;
    if (!out[key]) out[key] = [];
    out[key].push(record);
  });
  return out;
}

function core_domainsV2AuditCountBy_(records, getter) {
  var counts = {};
  (records || []).forEach(function(record) {
    var key = getter(record);
    if (!key) return;
    counts[key] = (counts[key] || 0) + 1;
  });
  return counts;
}

function core_domainsV2AuditReportDuplicates_(report, counts, code, message, limit) {
  limit = limit || 25;
  Object.keys(counts || {}).forEach(function(key) {
    if (counts[key] <= 1) return;
    core_domainsV2AuditIssue_(report, 'ERRO', code, message, {
      valor: key,
      ocorrencias: counts[key]
    });
  });
  if (report.totalErros > limit) {
    core_domainsV2AuditRecommendation_(report, 'Ha muitos erros de duplicidade; priorize saneamento por identificadores principais.');
  }
}

function coreAuditarPessoasV2_(options) {
  var report = core_domainsV2AuditNewReport_('PESSOAS_V2');
  var data = core_domainsV2AuditOpenDomain_('PESSOAS', report, options || {});

  var base = (data.PESSOAS_BASE && data.PESSOAS_BASE.records) || [];
  var identificadores = (data.PESSOAS_IDENTIFICADORES && data.PESSOAS_IDENTIFICADORES.records) || [];
  var detalhes = (data.MEMBROS_DETALHES && data.MEMBROS_DETALHES.records) || [];
  var vinculos = (data.VINCULOS_GEAPA && data.VINCULOS_GEAPA.records) || [];
  var resumo = (data.PESSOAS_RESUMO_OPERACIONAL && data.PESSOAS_RESUMO_OPERACIONAL.records) || [];

  var pessoasById = core_domainsV2AuditIndexBy_(base, 'ID_PESSOA');
  var detalhesByPessoa = core_domainsV2AuditIndexBy_(detalhes, 'ID_PESSOA');
  var resumoByPessoa = core_domainsV2AuditIndexBy_(resumo, 'ID_PESSOA');

  report.resumoQuantitativo = {
    pessoasBase: base.length,
    identificadores: identificadores.length,
    membrosDetalhes: detalhes.length,
    vinculos: vinculos.length,
    resumoOperacional: resumo.length
  };

  base.forEach(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    if (!idPessoa) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'PESSOA_SEM_ID_PESSOA', 'Registro em PESSOAS_BASE sem ID_PESSOA.', { row: record.__rowNumber });
    }
  });

  core_domainsV2AuditReportDuplicates_(
    report,
    core_domainsV2AuditCountBy_(base, function(record) { return String(record.ID_PESSOA || '').trim(); }),
    'ID_PESSOA_DUPLICADO',
    'ID_PESSOA duplicado em PESSOAS_BASE.'
  );
  core_domainsV2AuditReportDuplicates_(
    report,
    core_domainsV2AuditCountBy_(base, function(record) { return core_domainsV2AuditEmail_(record.EMAIL_PRINCIPAL); }),
    'EMAIL_DUPLICADO',
    'E-mail principal duplicado em PESSOAS_BASE.'
  );
  core_domainsV2AuditReportDuplicates_(
    report,
    core_domainsV2AuditCountBy_(base, function(record) { return core_domainsV2AuditDigits_(record.CPF); }),
    'CPF_DUPLICADO',
    'CPF duplicado em PESSOAS_BASE.'
  );
  core_domainsV2AuditReportDuplicates_(
    report,
    core_domainsV2AuditCountBy_(identificadores, function(record) {
      return core_domainsV2AuditStatus_(record.TIPO_IDENTIFICADOR) === 'RGA'
        ? String(record.VALOR_IDENTIFICADOR || '').trim()
        : '';
    }),
    'RGA_DUPLICADO',
    'RGA duplicado em PESSOAS_IDENTIFICADORES.'
  );

  identificadores.forEach(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    if (!idPessoa) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'IDENTIFICADOR_SEM_ID_PESSOA', 'Identificador sem ID_PESSOA.', { row: record.__rowNumber });
    } else if (!pessoasById[idPessoa]) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'IDENTIFICADOR_COM_PESSOA_INEXISTENTE', 'Identificador aponta para pessoa inexistente.', { idPessoa: idPessoa });
    }
  });

  base.forEach(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    if (idPessoa && !vinculos.some(function(vinculo) { return String(vinculo.ID_PESSOA || '').trim() === idPessoa; })) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'PESSOA_SEM_VINCULO', 'Pessoa sem vinculo registrado em VINCULOS_GEAPA.', { idPessoa: idPessoa });
    }
    if (idPessoa && !resumoByPessoa[idPessoa]) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'PESSOA_SEM_RESUMO_OPERACIONAL', 'Pessoa sem linha em PESSOAS_RESUMO_OPERACIONAL.', { idPessoa: idPessoa });
    }
  });

  vinculos.forEach(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    var tipo = core_domainsV2AuditTipoVinculo_(record.TIPO_VINCULO);
    var status = core_domainsV2AuditStatus_(record.STATUS_VINCULO);
    var ativo = core_domainsV2AuditIsSim_(record.ATIVO);

    if (!idPessoa || !pessoasById[idPessoa]) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'VINCULO_SEM_PESSOA', 'Vinculo sem pessoa existente em PESSOAS_BASE.', { idPessoa: idPessoa, row: record.__rowNumber });
    }
    if (tipo === 'MEMBRO_EFETIVO' && !detalhesByPessoa[idPessoa]) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'MEMBRO_EFETIVO_SEM_DETALHES', 'Membro efetivo sem MEMBROS_DETALHES.', { idPessoa: idPessoa });
    }
    if (tipo === 'MEMBRO_EM_ESPERA' && status && status !== 'EM_ESPERA') {
      core_domainsV2AuditIssue_(report, 'AVISO', 'MEMBRO_ESPERA_STATUS_INCOMPATIVEL', 'Membro em espera com status diferente de EM_ESPERA.', { idPessoa: idPessoa, status: record.STATUS_VINCULO });
    }
    if (tipo === 'EGRESSO' && !ativo) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'EGRESSO_INATIVO', 'Egresso com ATIVO diferente de SIM; revisar se representa estado atual da pessoa.', { idPessoa: idPessoa, status: record.STATUS_VINCULO, ativo: record.ATIVO });
    }
    if (tipo === 'MEMBRO_EFETIVO') {
      var hasRga = detalhesByPessoa[idPessoa] && detalhesByPessoa[idPessoa].some(function(detalhe) {
        return String(detalhe.RGA || '').trim();
      });
      if (!hasRga) {
        core_domainsV2AuditIssue_(report, 'ERRO', 'MEMBRO_EFETIVO_SEM_RGA', 'Membro efetivo sem RGA em MEMBROS_DETALHES.', { idPessoa: idPessoa });
      }
      var pessoaBase = pessoasById[idPessoa] ? pessoasById[idPessoa][0] : null;
      if ((status === 'ATIVO' || ativo) && pessoaBase && !core_domainsV2AuditEmail_(pessoaBase.EMAIL_PRINCIPAL)) {
        core_domainsV2AuditIssue_(report, 'AVISO', 'MEMBRO_EFETIVO_ATIVO_SEM_EMAIL', 'Membro efetivo ativo sem e-mail principal.', { idPessoa: idPessoa });
      }
      if ((status === 'ATIVO' || ativo) && pessoaBase && !core_domainsV2AuditIsSim_(pessoaBase.ATIVO)) {
        core_domainsV2AuditIssue_(report, 'AVISO', 'VINCULO_ATIVO_COM_PESSOA_INATIVA', 'Vinculo ativo aponta para pessoa com ATIVO diferente de SIM.', { idPessoa: idPessoa, ativoPessoa: pessoaBase.ATIVO });
      }
    }
  });

  resumo.forEach(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    if (idPessoa && !pessoasById[idPessoa]) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'RESUMO_COM_PESSOA_INEXISTENTE', 'Resumo operacional aponta para pessoa inexistente.', { idPessoa: idPessoa });
    }
  });

  if (report.totalErros || report.totalAvisos) {
    core_domainsV2AuditRecommendation_(report, 'Rodar auditoria apos saneamento e antes de qualquer recalculo real.');
  }
  if (!resumo.length) {
    core_domainsV2AuditRecommendation_(report, 'Rodar coreRecalcularPessoasResumoOperacionalV2({ dryRun: true }) antes de uso pelo portal.');
  }

  return report;
}

function coreAuditarVigenciasV2_(options) {
  var report = core_domainsV2AuditNewReport_('VIGENCIAS_V2');
  var vigenciasData = core_domainsV2AuditOpenDomain_('VIGENCIAS', report, options || {});
  var pessoasReportProbe = core_domainsV2AuditNewReport_('PESSOAS_V2_REFERENCIA');
  var pessoasData = core_domainsV2AuditOpenDomain_('PESSOAS', pessoasReportProbe, options || {});

  var funcoes = (vigenciasData.VIGENCIAS_FUNCOES && vigenciasData.VIGENCIAS_FUNCOES.records) || [];
  var cargos = (vigenciasData.CARGOS_CONFIG && vigenciasData.CARGOS_CONFIG.records) || [];
  var resumo = (vigenciasData.VIGENCIAS_RESUMO_ATUAL && vigenciasData.VIGENCIAS_RESUMO_ATUAL.records) || [];
  var pessoas = (pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [];
  var vinculos = (pessoasData.VINCULOS_GEAPA && pessoasData.VINCULOS_GEAPA.records) || [];

  var pessoasById = core_domainsV2AuditIndexBy_(pessoas, 'ID_PESSOA');
  var vinculosById = core_domainsV2AuditIndexBy_(vinculos, 'ID_VINCULO');
  var cargosByKey = core_domainsV2AuditIndexBy_(cargos, 'CARGO_KEY');

  report.resumoQuantitativo = {
    cargosConfig: cargos.length,
    vigenciasFuncoes: funcoes.length,
    vigenciasResumoAtual: resumo.length
  };

  cargos.forEach(function(record) {
    var cargoKey = String(record.CARGO_KEY || '').trim();
    var tipoFuncao = String(record.TIPO_FUNCAO || '').trim();
    if (!cargoKey) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'CARGO_CONFIG_SEM_CARGO_KEY', 'Cargo em CARGOS_CONFIG sem CARGO_KEY.', { row: record.__rowNumber });
    }
    if (!tipoFuncao) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'CARGO_CONFIG_SEM_TIPO_FUNCAO', 'Cargo em CARGOS_CONFIG sem TIPO_FUNCAO.', { cargoKey: cargoKey });
    }
    if (cargoKey && String(record.CARGO_NOME || '').trim() && cargoKey === String(record.CARGO_NOME || '').trim()) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'CARGO_KEY_PARECE_NOME_PUBLICO', 'CARGO_KEY parece usar nome publico em vez de chave padronizada.', { cargoKey: cargoKey });
    }
    if (core_domainsV2AuditIsSim_(record.GERA_PERFIL_PORTAL)) {
      var hasPermission = [
        'PODE_VER_AREA_DIRETORIA',
        'PODE_GERENCIAR_ATIVIDADES',
        'PODE_REGISTRAR_CHAMADA',
        'PODE_EDITAR_ATIVIDADE',
        'PODE_ANALISAR_JUSTIFICATIVAS',
        'PODE_GERENCIAR_CERTIFICADOS',
        'PODE_GERENCIAR_COMUNICACAO'
      ].some(function(field) {
        return core_domainsV2AuditIsSim_(record[field]);
      });
      if (!hasPermission && !String(record.PERFIL_PORTAL_PADRAO || '').trim()) {
        core_domainsV2AuditIssue_(report, 'AVISO', 'CARGO_COM_PERFIL_SEM_PERMISSOES', 'Cargo gera perfil de portal, mas nao tem permissoes/perfil padrao claros.', { cargoKey: cargoKey });
      }
    }
  });

  funcoes.forEach(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    var idVinculo = String(record.ID_VINCULO || '').trim();
    var cargoKey = String(record.CARGO_KEY || '').trim();
    var tipoFuncao = core_domainsV2AuditStatus_(record.TIPO_FUNCAO);
    var status = core_domainsV2AuditStatus_(record.STATUS_VIGENCIA);
    var dataFimReal = String(record.DATA_FIM_REAL || '').trim();
    var cargo = cargoKey && cargosByKey[cargoKey] ? cargosByKey[cargoKey][0] : null;

    if (!idPessoa) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'VIGENCIA_FUNCAO_SEM_ID_PESSOA', 'Vigencia de funcao sem ID_PESSOA.', { row: record.__rowNumber, cargoKey: cargoKey });
    } else if (!pessoasById[idPessoa]) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'VIGENCIA_FUNCAO_COM_PESSOA_INEXISTENTE', 'Vigencia aponta para pessoa inexistente.', { idPessoa: idPessoa, cargoKey: cargoKey });
    }
    if (!cargo) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'VIGENCIA_FUNCAO_CARGO_INVALIDO', 'Vigencia sem CARGO_KEY valido em CARGOS_CONFIG.', { cargoKey: cargoKey, row: record.__rowNumber });
    }
    var exigeVinculoAtivo = cargo && (core_domainsV2AuditIsSim_(cargo.EXIGE_MEMBRO_EFETIVO) || core_domainsV2AuditIsSim_(cargo.EXIGE_VINCULO_ATIVO));
    var vigenciaPareceAtiva = core_domainsV2AuditVigenciaPareceAtiva_(record);
    if (exigeVinculoAtivo && !idVinculo) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'VIGENCIA_EXIGE_MEMBRO_SEM_ID_VINCULO', 'Funcao exige membro efetivo, mas ID_VINCULO esta vazio.', { idPessoa: idPessoa, cargoKey: cargoKey });
    }
    if (idVinculo && !vinculosById[idVinculo]) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'VIGENCIA_COM_VINCULO_INEXISTENTE', 'Funcao aponta para ID_VINCULO inexistente em Pessoas v2.', { idVinculo: idVinculo, idPessoa: idPessoa });
    } else if (exigeVinculoAtivo && vigenciaPareceAtiva && idVinculo && vinculosById[idVinculo]) {
      var vinculo = vinculosById[idVinculo][0];
      if (core_domainsV2AuditStatus_(vinculo.STATUS_VINCULO) !== 'ATIVO' && !core_domainsV2AuditIsSim_(vinculo.ATIVO)) {
        core_domainsV2AuditIssue_(report, 'AVISO', 'VIGENCIA_EXIGE_MEMBRO_COM_VINCULO_INATIVO', 'Funcao exige vinculo ativo, mas o vinculo nao parece ativo.', {
          idVinculo: idVinculo,
          idPessoa: idPessoa,
          statusVinculo: vinculo.STATUS_VINCULO,
          ativoVinculo: vinculo.ATIVO
        });
      }
    }
    if ((tipoFuncao === 'DIRETORIA' || tipoFuncao === 'ASSESSORIA') && !String(record.ID_DIRETORIA || '').trim()) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'VIGENCIA_SEM_DIRETORIA', 'Funcao de diretoria/assessoria sem ID_DIRETORIA.', { idPessoa: idPessoa, cargoKey: cargoKey });
    }
    if ((status === 'ATIVA' || status === 'ATIVO') && dataFimReal) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'VIGENCIA_ATIVA_COM_DATA_FIM_REAL', 'Vigencia ativa com DATA_FIM_REAL preenchida.', { idPessoa: idPessoa, cargoKey: cargoKey });
    }
    if ((status === 'ENCERRADA' || status === 'INATIVA') && !dataFimReal) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'VIGENCIA_ENCERRADA_SEM_DATA_FIM_REAL', 'Vigencia encerrada sem DATA_FIM_REAL.', { idPessoa: idPessoa, cargoKey: cargoKey });
    }
    if (tipoFuncao === 'CONSELHEIRO' || tipoFuncao === 'ASSESSOR') {
      core_domainsV2AuditIssue_(report, 'AVISO', 'TIPO_FUNCAO_PARECE_CATEGORIA_PESSOA', 'Tipo de funcao parece categoria de pessoa; revisar se deve ser CONSELHO/ASSESSORIA.', { tipoFuncao: record.TIPO_FUNCAO, cargoKey: cargoKey });
    }
  });

  if (!resumo.length) {
    core_domainsV2AuditIssue_(report, 'AVISO', 'VIGENCIAS_RESUMO_ATUAL_VAZIO', 'VIGENCIAS_RESUMO_ATUAL esta vazio ou sem registros.');
    core_domainsV2AuditRecommendation_(report, 'Rodar coreRecalcularVigenciasResumoAtualV2({ dryRun: true }) para conferir perfis e permissoes calculadas antes do uso pelo portal.');
  }
  if (pessoasReportProbe.totalErros) {
    core_domainsV2AuditIssue_(report, 'AVISO', 'PESSOAS_REFERENCIA_COM_ERROS', 'Auditoria de Vigencias depende de Pessoas v2, que tem erros estruturais.', { totalErrosPessoas: pessoasReportProbe.totalErros });
  }

  return report;
}

function core_domainsV2ResumoOptions_(options) {
  options = options || {};
  return {
    dryRun: options.dryRun !== false,
    confirmacao: String(options.confirmacao || '').trim()
  };
}

function core_domainsV2ResumoPushDistinct_(target, value) {
  var text = String(value || '').trim();
  if (text && target.indexOf(text) < 0) target.push(text);
}

function core_domainsV2ResumoTimestamp_(value) {
  if (!value) return 0;
  if (Object.prototype.toString.call(value) === '[object Date]') return value.getTime();
  var parsed = Date.parse(String(value));
  return isNaN(parsed) ? 0 : parsed;
}

function core_domainsV2ResumoVigenciaAtual_(record) {
  var status = core_domainsV2AuditStatus_(record.STATUS_VIGENCIA);
  var dataFimReal = String(record.DATA_FIM_REAL || '').trim();
  if (dataFimReal) return false;
  if (status === 'ATIVA' || status === 'ATIVO') return true;
  if (status === 'ENCERRADA' || status === 'ENCERRADO' || status === 'INATIVA' || status === 'INATIVO') return false;
  return core_domainsV2AuditIsSim_(record.ATIVO);
}

function core_domainsV2ResumoPermissionsFromCargo_(cargo) {
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
    if (core_domainsV2AuditIsSim_(cargo[field])) permissions.push(field);
  });
  return permissions;
}

function core_domainsV2ResumoBuildRows_(vigenciasData, pessoasData) {
  var pessoasById = core_domainsV2AuditIndexBy_(pessoasData.PESSOAS_BASE.records || [], 'ID_PESSOA');
  var detalhesByPessoa = core_domainsV2AuditIndexBy_(pessoasData.MEMBROS_DETALHES.records || [], 'ID_PESSOA');
  var cargosByKey = core_domainsV2AuditIndexBy_(vigenciasData.CARGOS_CONFIG.records || [], 'CARGO_KEY');
  var groups = {};

  (vigenciasData.VIGENCIAS_FUNCOES.records || []).forEach(function(vigencia) {
    if (!core_domainsV2ResumoVigenciaAtual_(vigencia)) return;

    var idPessoa = String(vigencia.ID_PESSOA || '').trim();
    if (!idPessoa) return;

    var cargoKey = String(vigencia.CARGO_KEY || '').trim();
    var cargo = cargoKey && cargosByKey[cargoKey] ? cargosByKey[cargoKey][0] : {};
    var group = groups[idPessoa];
    if (!group) {
      group = {
        idPessoa: idPessoa,
        vigencias: [],
        cargos: [],
        tipos: [],
        grupos: [],
        diretorias: [],
        perfis: [],
        permissions: []
      };
      groups[idPessoa] = group;
    }

    group.vigencias.push(vigencia);
    core_domainsV2ResumoPushDistinct_(group.cargos, vigencia.CARGO_NOME_SNAPSHOT || cargo.NOME_PUBLICO || cargo.CARGO_NOME || cargoKey);
    core_domainsV2ResumoPushDistinct_(group.tipos, vigencia.TIPO_FUNCAO || cargo.TIPO_FUNCAO);
    core_domainsV2ResumoPushDistinct_(group.grupos, cargo.GRUPO_FUNCAO || cargo.GRUPO_CARGO);
    core_domainsV2ResumoPushDistinct_(group.diretorias, vigencia.ID_DIRETORIA);
    if (core_domainsV2AuditIsSim_(cargo.GERA_PERFIL_PORTAL)) {
      core_domainsV2ResumoPushDistinct_(group.perfis, cargo.PERFIL_PORTAL_PADRAO || cargo.CARGO_KEY);
    }
    core_domainsV2ResumoPermissionsFromCargo_(cargo).forEach(function(permission) {
      core_domainsV2ResumoPushDistinct_(group.permissions, permission);
    });
  });

  return Object.keys(groups).map(function(idPessoa) {
    var group = groups[idPessoa];
    group.vigencias.sort(function(a, b) {
      return core_domainsV2ResumoTimestamp_(b.DATA_INICIO) - core_domainsV2ResumoTimestamp_(a.DATA_INICIO);
    });
    var main = group.vigencias[0] || {};
    var pessoa = pessoasById[idPessoa] ? pessoasById[idPessoa][0] : {};
    var detalhes = detalhesByPessoa[idPessoa] ? detalhesByPessoa[idPessoa][0] : {};

    return {
      ID_PESSOA: idPessoa,
      NOME_EXIBICAO: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || '',
      RGA: detalhes.RGA || '',
      CARGO_FUNCAO_ATUAL: group.cargos.join('; '),
      TIPO_FUNCAO_ATUAL: group.tipos.join('; '),
      GRUPO_FUNCAO_ATUAL: group.grupos.join('; '),
      ID_DIRETORIA_ATUAL: group.diretorias.join('; '),
      PERFIS_PORTAL_CALCULADOS: group.perfis.join('; '),
      PERMISSOES_CALCULADAS: group.permissions.join('; '),
      DATA_INICIO_FUNCAO_ATUAL: main.DATA_INICIO || '',
      DATA_FIM_PREVISTA: main.DATA_FIM_PREVISTA || '',
      ULTIMA_ATUALIZACAO: new Date()
    };
  }).sort(function(a, b) {
    return String(a.NOME_EXIBICAO || a.ID_PESSOA).localeCompare(String(b.NOME_EXIBICAO || b.ID_PESSOA));
  });
}

function coreRecalcularVigenciasResumoAtualV2_(options) {
  var opts = core_domainsV2ResumoOptions_(options);
  if (!opts.dryRun && opts.confirmacao !== 'RECALCULAR_RESUMO_ATUAL_V2') {
    throw new Error('Para escrever VIGENCIAS_RESUMO_ATUAL, informe confirmacao: "RECALCULAR_RESUMO_ATUAL_V2".');
  }

  var report = core_domainsV2AuditNewReport_('RECALCULAR_VIGENCIAS_RESUMO_ATUAL_V2');
  report.dryRun = opts.dryRun;

  var vigenciasData = core_domainsV2AuditOpenDomain_('VIGENCIAS', report);
  var pessoasData = core_domainsV2AuditOpenDomain_('PESSOAS', report);
  if (report.totalErros > 0) return report;

  var resumoSheet = vigenciasData.VIGENCIAS_RESUMO_ATUAL.sheet;
  var headers = vigenciasData.VIGENCIAS_RESUMO_ATUAL.headers || [];
  var rows = core_domainsV2ResumoBuildRows_(vigenciasData, pessoasData);

  report.resumoQuantitativo = {
    vigenciasFuncoes: (vigenciasData.VIGENCIAS_FUNCOES.records || []).length,
    linhasCalculadas: rows.length,
    linhasExistentesAntes: (vigenciasData.VIGENCIAS_RESUMO_ATUAL.records || []).length
  };

  if (opts.dryRun) {
    report.amostra = rows.slice(0, 10);
    core_domainsV2AuditRecommendation_(report, 'Conferir a amostra; se estiver correta, rodar coreRecalcularVigenciasResumoAtualV2({ dryRun: false, confirmacao: "RECALCULAR_RESUMO_ATUAL_V2" }).');
    return report;
  }

  if (resumoSheet.getLastRow() > 1) {
    resumoSheet.getRange(2, 1, resumoSheet.getLastRow() - 1, resumoSheet.getLastColumn()).clearContent();
  }
  if (rows.length > 0) {
    var values = rows.map(function(row) {
      return core_buildRowFromObjectByHeaders_(headers, row);
    });
    resumoSheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }

  report.resumoQuantitativo.linhasEscritas = rows.length;
  return report;
}

function coreAuditarDominiosCentraisV2_(options) {
  var pessoas = coreAuditarPessoasV2_(options || {});
  var vigencias = coreAuditarVigenciasV2_(options || {});
  var report = core_domainsV2AuditNewReport_('DOMINIOS_CENTRAIS_V2');

  report.erros = [].concat(pessoas.erros, vigencias.erros);
  report.avisos = [].concat(pessoas.avisos, vigencias.avisos);
  report.totalErros = report.erros.length;
  report.totalAvisos = report.avisos.length;
  report.ok = report.totalErros === 0;
  report.recomendacoes = [].concat(pessoas.recomendacoes, vigencias.recomendacoes).filter(function(value, index, arr) {
    return arr.indexOf(value) === index;
  });
  report.resumoQuantitativo = {
    pessoas: pessoas.resumoQuantitativo,
    vigencias: vigencias.resumoQuantitativo
  };
  report.pessoas = pessoas;
  report.vigencias = vigencias;

  core_domainsV2AuditIssue_(report, 'AVISO', 'ATIVIDADES_V2_NAO_AUDITADA_NESTA_FASE', 'Checagens cruzadas com Atividades v2 ainda nao foram executadas nesta primeira auditoria central.');
  core_domainsV2AuditRecommendation_(report, 'Checagens com Atividades v2 ficam pendentes ate os contratos de Atividades v2 estarem estabilizados no Registry.');
  return report;
}
