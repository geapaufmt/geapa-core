'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const domain = { console, Object, Array, String, Number, Boolean, Error, Date, JSON, Math, RegExp, isNaN };
vm.createContext(domain);
['01_core_sheets.js', '30_core_domains_v2_audit.js', '31_core_domains_v2_operational_api.js', '33_core_v2_pessoas_vigencias_rotinas.js']
  .forEach((file) => vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), domain, { filename: file }));

const portal = { console, Object, Array, String, Number, Boolean, Error, Date, JSON, Math, RegExp, isNaN };
vm.createContext(portal);
['01_core_sheets.js', '30_core_domains_v2_audit.js', '31_core_domains_v2_operational_api.js']
  .forEach((file) => vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), portal, { filename: file }));
vm.runInContext(fs.readFileSync(path.join(root, '26_core_portal_access.js'), 'utf8'), portal, { filename: '26_core_portal_access.js' });

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test('normaliza aliases de membro ingressante nas duas rotas do Core', () => {
  assert.equal(domain.core_domainsV2AuditTipoVinculo_('ingressante'), 'MEMBRO_INGRESSANTE');
  assert.equal(domain.core_domainsV2NormalizeTipoVinculo_('membro ingressante'), 'MEMBRO_INGRESSANTE');
  assert.equal(domain.core_v2RotinasTipoVinculo_('ingressante'), 'MEMBRO_INGRESSANTE');
});

test('vinculo efetivo tem prioridade e ingressante precede membro em espera', () => {
  const active = (tipo) => ({ TIPO_VINCULO: tipo, STATUS_VINCULO: 'ATIVO', ATIVO: 'SIM', DATA_INICIO: '2026-08-21' });
  assert.equal(domain.core_domainsV2PickCurrentVinculo_([active('MEMBRO_EM_ESPERA'), active('MEMBRO_INGRESSANTE')]).TIPO_VINCULO, 'MEMBRO_INGRESSANTE');
  assert.equal(domain.core_domainsV2PickCurrentVinculo_([active('MEMBRO_INGRESSANTE'), active('MEMBRO_EFETIVO')]).TIPO_VINCULO, 'MEMBRO_EFETIVO');
});

test('ingressante ativo recebe apenas perfil basico mesmo com excecao permissiva', () => {
  const state = domain.core_domainsV2BuildPortalState_(
    { TIPO_VINCULO: 'MEMBRO_INGRESSANTE', STATUS_VINCULO: 'ATIVO', ATIVO: 'SIM' },
    { PERFIS_PORTAL_CALCULADOS: 'DIRETORIA' },
    [{ STATUS: 'ATIVO', PERFIL_EXTRA: 'ADMIN' }]
  );
  assert.deepEqual(JSON.parse(JSON.stringify(state)), { perfil: 'MEMBRO_INGRESSANTE', portalAtivo: 'SIM' });
});

test('ingressante fica inelegivel para diretoria e conta como membro atual, nao efetivo', () => {
  const link = { TIPO_VINCULO: 'MEMBRO_INGRESSANTE', STATUS_VINCULO: 'ATIVO', ATIVO: 'SIM' };
  assert.equal(domain.core_domainsV2Eligibility_(link), 'NAO_ELEGIVEL_INGRESSANTE');
  const stats = domain.core_domainsV2PessoasResumoStats_([{ TIPO_VINCULO_ATUAL: 'MEMBRO_INGRESSANTE', STATUS_VINCULO_ATUAL: 'ATIVO', PORTAL_ATIVO: 'SIM' }]);
  assert.equal(stats.totalMembrosAtivos, 1);
  assert.equal(stats.totalMembrosIngressantes, 1);
  assert.equal(stats.totalMembrosEfetivos, 0);
});

test('contrato do Portal libera acesso basico sem permissoes de gestao', () => {
  assert.equal(portal.corePortalLinkAllowsMemberAccess_({ tipoVinculo: 'MEMBRO_INGRESSANTE', statusVinculo: 'ATIVO' }), true);
  const permissions = vm.runInContext('Array.from(CORE_PORTAL_V2_MIN_PERMISSIONS.MEMBRO_INGRESSANTE)', portal);
  assert.deepEqual(Array.from(permissions), ['portal:acessar', 'situacao:ver_propria']);
  assert.equal(permissions.some((item) => /admin|gest|diret|membros:|atividades:/.test(item)), false);
});

test('excecao permissiva nao promove ingressante a perfil ou permissao de gestao', () => {
  const bundle = {
    pessoa: { ID_PESSOA: 'PES-INGRESSANTE' },
    resumoOperacional: { TIPO_VINCULO_ATUAL: 'MEMBRO_INGRESSANTE', STATUS_VINCULO_ATUAL: 'ATIVO', PERFIL_PORTAL_CALCULADO: 'DIRETORIA', PORTAL_ATIVO: 'SIM' },
    vinculos: [{ ID_PESSOA: 'PES-INGRESSANTE', TIPO_VINCULO: 'MEMBRO_INGRESSANTE', STATUS_VINCULO: 'ATIVO', ATIVO: 'SIM' }],
    portalExcecoes: [{ STATUS: 'ATIVO', ATIVO: 'SIM', PERFIL_EXTRA: 'ADMIN', PERMISSAO_EXTRA: 'sistema:admin; atividades:gerir' }]
  };
  const profile = portal.corePortalCalcularPerfilEfetivo_('PES-INGRESSANTE', {
    bundle,
    profileMap: { MEMBRO_INGRESSANTE: { PERFIL_PORTAL: 'MEMBRO_INGRESSANTE' } },
    vigenciasResumo: {}
  });
  assert.equal(profile.perfilPortalEfetivo, 'MEMBRO_INGRESSANTE');
  assert.deepEqual(Array.from(profile.perfisPortal), ['MEMBRO_INGRESSANTE']);
  const permissions = portal.corePortalListarPermissoesEfetivas_('PES-INGRESSANTE', { bundle, profileResult: profile });
  assert.deepEqual(Array.from(permissions.permissoes), ['portal:acessar', 'situacao:ver_propria']);
});

test('API publica separa lista de membros atuais da lista apenas de efetivos', () => {
  const operational = fs.readFileSync(path.join(root, '31_core_domains_v2_operational_api.js'), 'utf8');
  const exportsSource = fs.readFileSync(path.join(root, '20_public_exports.js'), 'utf8');
  assert.match(operational, /corePessoasListCurrentMembers_[\s\S]*MEMBRO_INGRESSANTE[\s\S]*MEMBRO_EFETIVO/);
  assert.match(operational, /function corePessoasListEffectiveMembers_/);
  assert.match(exportsSource, /function corePessoasListEffectiveMembers/);
});

test('busca administrativa reconhece ID e as novas categorias de membro', () => {
  const item = domain.core_pessoasAdminPortalMapRow_({
    ID_PESSOA: 'PES-000096', NOME_EXIBICAO: 'Pessoa Ingressante', RGA: '20260096',
    EMAIL: 'ingressante@example.invalid', TIPO_VINCULO_ATUAL: 'MEMBRO_INGRESSANTE',
    STATUS_VINCULO_ATUAL: 'ATIVO', PERFIL_PORTAL_CALCULADO: 'MEMBRO_INGRESSANTE',
    FREQUENCIA_RESUMIDA: ''
  });
  assert.equal(domain.core_pessoasAdminPortalMatches_(item, domain.core_pessoasAdminPortalFilters_({ texto: 'PES-000096' })), true);
  assert.equal(domain.core_pessoasAdminPortalMatches_(item, domain.core_pessoasAdminPortalFilters_({ tipoVinculo: 'MEMBRO_INGRESSANTE' })), true);
  assert.equal(item.situacaoFrequencia, 'SEM_DADOS');
});

test('filtro de perfil encontra perfil dentro de combinacao e frequencia usa categorias', () => {
  const item = domain.core_pessoasAdminPortalMapRow_({
    ID_PESSOA: 'PES-000095', TIPO_VINCULO_ATUAL: 'MEMBRO_EFETIVO', STATUS_VINCULO_ATUAL: 'ATIVO',
    PERFIL_PORTAL_CALCULADO: 'DIRETORIA; ADMIN', FREQUENCIA_RESUMIDA: 'Frequencia 80%; Situacao REGULAR'
  });
  assert.equal(domain.core_pessoasAdminPortalMatches_(item, domain.core_pessoasAdminPortalFilters_({ perfilPortal: 'DIRETORIA' })), true);
  assert.equal(domain.core_pessoasAdminPortalMatches_(item, domain.core_pessoasAdminPortalFilters_({ situacaoFrequencia: 'REGULAR' })), true);
  const options = domain.core_pessoasAdminPortalFilterOptions_([item]);
  assert.equal(Array.from(options.perfisPortal).includes('DIRETORIA'), true);
  assert.equal(Array.from(options.perfisPortal).includes('DIRETORIA; ADMIN'), false);
  assert.equal(Array.from(options.tiposVinculo).includes('MEMBRO_INGRESSANTE'), true);
  assert.deepEqual(Array.from(options.frequencias), ['REGULAR', 'COM_FALTAS', 'SEM_DADOS']);
});

let failures = 0;
for (const item of tests) {
  try { item.fn(); process.stdout.write(`ok - ${item.name}\n`); }
  catch (error) { failures += 1; process.stderr.write(`not ok - ${item.name}\n${error.stack}\n`); }
}
if (failures) process.exitCode = 1;
