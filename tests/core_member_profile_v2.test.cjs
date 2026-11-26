const assert = require('node:assert/strict');
const {test} = require('node:test');
const {isolatedDomain, enableCompatibility} = require('./canonical-domain.cjs');
const domain = isolatedDomain();
const {context} = domain;
enableCompatibility(domain);

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

test('setup legado de membros permanece fechado em DEV e PROD', () => {
  const current = isolatedDomain().context;
  for (const ambiente of ['DEV', 'PROD']) {
    assert.throws(() => current.core_setupMemberEvolutionV2_({ambiente}), {code:'MEMBERS_LEGACY_FENCED'});
  }
});

test('mapa de dominio inclui as quatro fontes novas e schemas aditivos', () => {
  assert.equal(context.CORE_DOMAIN_V2_MAP.PESSOAS.sheets.CURSOS_CATALOGO, 'CURSOS_CATALOGO');
  assert.equal(context.CORE_DOMAIN_V2_MAP.PESSOAS.specificRegistryKeys.INGRESSOS_MEMBROS, 'PESSOAS_V2_INGRESSOS_MEMBROS');
  const names = context.CORE_DOMAINS_V2_SCHEMAS.PESSOAS.map((item) => item.sheetName);
  assert.ok(names.includes('RESPOSTAS_AVALIACAO_EGRESSOS'));
});
