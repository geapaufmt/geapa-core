const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.join(__dirname, '..');
const context = {
  Object, Array, String, Number, Boolean, Error, Date, JSON, Math, RegExp,
  isNaN, console,
  core_parseEntrySemesterFromRga_: undefined,
  core_openDomainSpreadsheet_: () => { throw new Error('PLANILHA_REAL_PROIBIDA'); },
  core_withLock_: () => { throw new Error('ESCRITA_PROIBIDA_EM_TESTE'); }
};
vm.createContext(context);
[
  '39a_core_locality_catalog_data.js',
  '28_core_domains_v2_setup.js',
  '28a_core_domains_v2_resolver.js',
  '39_core_member_profile_v2.js'
].forEach((file) => vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), context, { filename: file }));

function sheet(name, headers, records = []) {
  const rows = [headers, ...records.map((record) => headers.map((header) => record[header] ?? ''))];
  return {
    getName: () => name,
    getLastColumn: () => headers.length,
    getLastRow: () => rows.length,
    getRange: (row, column, rowCount, columnCount) => ({
      getValues: () => rows.slice(row - 1, row - 1 + rowCount)
        .map((values) => values.slice(column - 1, column - 1 + columnCount))
    })
  };
}

function workbook(overrides = {}) {
  const defaults = {
    PESSOAS_BASE: sheet('PESSOAS_BASE', ['ID_PESSOA']),
    MEMBROS_DETALHES: sheet('MEMBROS_DETALHES', ['ID_PESSOA']),
    CURSOS_CATALOGO: sheet('CURSOS_CATALOGO', ['CURSO_ID']),
    INGRESSOS_MEMBROS: null,
    CONVITES_AVALIACAO_EGRESSOS: null,
    RESPOSTAS_AVALIACAO_EGRESSOS: null
  };
  const sheets = Object.assign(defaults, overrides);
  return { getId: () => 'fake-dev-spreadsheet', getSheetByName: (name) => sheets[name] || null };
}

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test('catalogo oficial versionado contem paises, 27 UFs e municipios IBGE', () => {
  assert.equal(context.CORE_LOCALIDADES_CATALOGO_META.fonte, 'IBGE API de Localidades');
  assert.equal(context.CORE_LOCALIDADES_UFS.length, 27);
  assert.equal(context.CORE_LOCALIDADES_CATALOGO_META.quantidadeMunicipios, 5571);
  assert.equal(context.CORE_LOCALIDADES_PAISES.find((item) => item.codigo === 'BR').nome, 'Brasil');
  assert.equal(context.core_findMunicipalityV2_('MT', '5103403').nome, 'Cuiabá');
});

test('origem brasileira valida codigo, nome e UF de forma atomica', () => {
  const result = context.core_validateOriginV2_({
    paisOrigemCodigo: 'BR', ufOrigem: 'MT', municipioOrigemCodigo: '5103403', cidadeNatal: 'Cuiabá'
  });
  assert.equal(result.PAIS_ORIGEM_NOME, 'Brasil');
  assert.equal(result.UF_ORIGEM, 'MT');
  assert.equal(result.MUNICIPIO_ORIGEM_CODIGO, '5103403');
  assert.equal(result.CIDADE_NATAL, 'Cuiabá');
});

test('origem brasileira rejeita municipio incompatível e nome divergente', () => {
  assert.throws(() => context.core_validateOriginV2_({
    paisOrigemCodigo: 'BR', ufOrigem: 'GO', municipioOrigemCodigo: '5103403'
  }), (error) => error.code === 'MUNICIPIO_ORIGEM_INVALIDO');
  assert.throws(() => context.core_validateOriginV2_({
    paisOrigemCodigo: 'BR', ufOrigem: 'MT', municipioOrigemCodigo: '5103403', cidadeNatal: 'Sinop'
  }), (error) => error.code === 'MUNICIPIO_ORIGEM_NOME_DIVERGENTE');
});

test('origem estrangeira aceita cidade e regiao, mas rejeita UF', () => {
  const valid = context.core_validateOriginV2_({ paisOrigemCodigo: 'PY', cidadeNatal: 'Assunção', regiaoOrigem: 'Central' });
  assert.equal(valid.UF_ORIGEM, '');
  assert.equal(valid.CIDADE_NATAL, 'Assunção');
  assert.throws(() => context.core_validateOriginV2_({ paisOrigemCodigo: 'PY', cidadeNatal: 'Assunção', ufOrigem: 'MT' }),
    (error) => error.code === 'LOCALIDADE_ESTRANGEIRA_UF_NAO_PERMITIDA');
});

test('curso precisa existir, estar ativo e OUTRO exige descricao', () => {
  const rows = [
    { CURSO_ID: 'AGRONOMIA_UFMT_SINOP', NOME_CURSO: 'Agronomia', INSTITUICAO: 'UFMT', CAMPUS: 'Sinop', NIVEL: 'GRADUACAO', ATIVO: 'SIM', PERMITE_CADASTRO: 'SIM' },
    { CURSO_ID: 'INATIVO', NOME_CURSO: 'Inativo', ATIVO: 'NAO', PERMITE_CADASTRO: 'SIM' },
    { CURSO_ID: 'OUTRO', NOME_CURSO: 'Outro', ATIVO: 'SIM', PERMITE_CADASTRO: 'SIM' }
  ];
  assert.equal(context.core_validateCourseV2_({ cursoId: 'AGRONOMIA_UFMT_SINOP' }, rows).CAMPUS, 'Sinop');
  assert.throws(() => context.core_validateCourseV2_({ cursoId: 'NAO_EXISTE' }, rows), (error) => error.code === 'CURSO_ID_INVALIDO');
  assert.throws(() => context.core_validateCourseV2_({ cursoId: 'INATIVO' }, rows), (error) => error.code === 'CURSO_INATIVO');
  assert.throws(() => context.core_validateCourseV2_({ cursoId: 'OUTRO' }, rows), (error) => error.code === 'CURSO_OUTRO_EXIGE_DESCRICAO');
});

const semesters = [
  { ID_SEMESTRE: '2023/1', ANO: 2023, SEMESTRE: 1, DATA_INICIO: '2023-03-01', DATA_FIM: '2023-07-15', STATUS: 'ENCERRADO' },
  { ID_SEMESTRE: '2023/2', ANO: 2023, SEMESTRE: 2, DATA_INICIO: '2023-08-15', DATA_FIM: '2023-12-20', STATUS: 'ENCERRADO' },
  { ID_SEMESTRE: '2024/1', ANO: 2024, SEMESTRE: 1, DATA_INICIO: '2024-04-01', DATA_FIM: '2024-08-10', STATUS: 'ATIVO' },
  { ID_SEMESTRE: '2024/2', ANO: 2024, SEMESTRE: 2, DATA_INICIO: '2024-09-01', DATA_FIM: '2025-01-10', STATUS: 'PLANEJADO' }
];

test('periodo vem do RGA e semestre ordinal conta a sequencia institucional', () => {
  const first = context.core_calculateAcademicSemesterV2_('20231ABC', semesters, '2023-05-01');
  const second = context.core_calculateAcademicSemesterV2_('20232ABC', semesters, '2024-05-01');
  assert.equal(first.periodoIngressoCurso, '2023/1');
  assert.equal(first.semestreAtualCalculado, 1);
  assert.equal(second.periodoIngressoCurso, '2023/2');
  assert.equal(second.semestreAtualCalculado, 2);
});

test('intervalo e semestre futuro planejado nao incrementam o calculo', () => {
  assert.equal(context.core_calculateAcademicSemesterV2_('20231ABC', semesters, '2024-02-01').semestreAtualCalculado, 2);
  assert.equal(context.core_calculateAcademicSemesterV2_('20231ABC', semesters, '2024-08-20').semestreAtualCalculado, 3);
  assert.throws(() => context.core_calculateAcademicSemesterV2_('RGA-INVALIDO', semesters, '2024-05-01'),
    (error) => error.code === 'RGA_PERIODO_INGRESSO_INVALIDO');
});

test('completude usa regra deterministica sem inventar percentual', () => {
  const result = context.core_calculateProfileCompletenessV2_({ NOME_COMPLETO: 'A' }, { RGA: '20231' });
  assert.equal(result.status, 'PARCIAL');
  assert.equal('percentual' in result, false);
});

test('setup padrao e dry-run, nao abre planilha real nem escreve', () => {
  const result = context.core_setupMemberEvolutionV2_({ spreadsheet: workbook(), spreadsheetId: 'fake-dev-spreadsheet' });
  assert.equal(result.dryRun, true);
  assert.equal(result.escritaExecutada, false);
  assert.equal(result.idempotente, true);
  assert.equal(result.tokenConfirmacao, 'PREPARAR_EVOLUCAO_MEMBROS_V2_DEV');
});

test('segundo dry-run nao recomenda cabecalhos ou cursos ja existentes', () => {
  const courseHeaders = Array.from(context.CORE_MEMBER_EVOLUTION_SCHEMAS.CURSOS_CATALOGO);
  const result = context.core_setupMemberEvolutionV2_({ spreadsheet: workbook({
    PESSOAS_BASE: sheet('PESSOAS_BASE', Array.from(context.CORE_MEMBER_EVOLUTION_SCHEMAS.PESSOAS_BASE)),
    MEMBROS_DETALHES: sheet('MEMBROS_DETALHES', Array.from(context.CORE_MEMBER_EVOLUTION_SCHEMAS.MEMBROS_DETALHES)),
    CURSOS_CATALOGO: sheet('CURSOS_CATALOGO', courseHeaders, [
      { CURSO_ID: 'AGRONOMIA_UFMT_SINOP' }, { CURSO_ID: 'OUTRO' }
    ]),
    INGRESSOS_MEMBROS: sheet('INGRESSOS_MEMBROS', Array.from(context.CORE_MEMBER_EVOLUTION_SCHEMAS.INGRESSOS_MEMBROS)),
    CONVITES_AVALIACAO_EGRESSOS: sheet('CONVITES_AVALIACAO_EGRESSOS', Array.from(context.CORE_MEMBER_EVOLUTION_SCHEMAS.CONVITES_AVALIACAO_EGRESSOS)),
    RESPOSTAS_AVALIACAO_EGRESSOS: sheet('RESPOSTAS_AVALIACAO_EGRESSOS', Array.from(context.CORE_MEMBER_EVOLUTION_SCHEMAS.RESPOSTAS_AVALIACAO_EGRESSOS))
  }) });
  result.abas.forEach((plan) => assert.equal(plan.cabecalhosAusentes.length, 0));
  assert.equal(result.abas.find((plan) => plan.aba === 'CURSOS_CATALOGO').linhasRecomendadas.length, 0);
});

test('setup recusa PROD inclusive em dry-run', () => {
  assert.throws(() => context.core_setupMemberEvolutionV2_({ ambiente: 'PROD', spreadsheet: workbook() }),
    (error) => error.code === 'SETUP_EVOLUCAO_MEMBROS_PROD_PROIBIDO');
});

test('mapa de dominio inclui as quatro fontes novas e schemas aditivos', () => {
  assert.equal(context.CORE_DOMAIN_V2_MAP.PESSOAS.sheets.CURSOS_CATALOGO, 'CURSOS_CATALOGO');
  assert.equal(context.CORE_DOMAIN_V2_MAP.PESSOAS.specificRegistryKeys.INGRESSOS_MEMBROS, 'PESSOAS_V2_INGRESSOS_MEMBROS');
  const names = context.CORE_DOMAINS_V2_SCHEMAS.PESSOAS.map((item) => item.sheetName);
  assert.ok(names.includes('RESPOSTAS_AVALIACAO_EGRESSOS'));
});

let failures = 0;
for (const item of tests) {
  try { item.fn(); process.stdout.write(`ok - ${item.name}\n`); }
  catch (error) { failures += 1; process.stderr.write(`not ok - ${item.name}\n${error.stack}\n`); }
}
if (failures) process.exitCode = 1;
