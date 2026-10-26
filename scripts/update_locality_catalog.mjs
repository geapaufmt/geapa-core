import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';

const root = path.resolve(import.meta.dirname, '..');
const source = 'https://servicodados.ibge.gov.br/api/docs/localidades';
const apiVersion = '1.0.0';
const catalogDate = new Date().toISOString().slice(0, 10);

async function readJson(url) {
  const response = await fetch(url, { headers: { accept: 'application/json' } });
  if (!response.ok) throw new Error(`IBGE_HTTP_${response.status}:${url}`);
  return response.json();
}

function normalizeOfficialText(value) {
  const text = String(value || '').trim();
  // O endpoint de paises ocasionalmente declara UTF-8 para bytes interpretados
  // como latin1. Corrige somente a assinatura inequivoca de mojibake.
  return /(?:Ã.|Â.)/.test(text) ? Buffer.from(text, 'latin1').toString('utf8') : text;
}

const [countriesRaw, statesRaw, municipalitiesRaw] = await Promise.all([
  readJson('https://servicodados.ibge.gov.br/api/v1/localidades/paises'),
  readJson('https://servicodados.ibge.gov.br/api/v1/localidades/estados?orderBy=nome'),
  readJson('https://servicodados.ibge.gov.br/api/v1/localidades/municipios?orderBy=nome')
]);

const countries = countriesRaw
  .map((item) => ({ codigo: String(item?.id?.['ISO-ALPHA-2'] || '').toUpperCase(), nome: normalizeOfficialText(item?.nome) }))
  .filter((item) => /^[A-Z]{2}$/.test(item.codigo) && item.nome)
  .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));

const states = statesRaw
  .map((item) => ({ codigo: Number(item.id), sigla: String(item.sigla || '').toUpperCase(), nome: normalizeOfficialText(item.nome) }))
  .filter((item) => item.codigo && /^[A-Z]{2}$/.test(item.sigla) && item.nome)
  .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));

const municipalities = municipalitiesRaw.map((item) => {
  const uf = item?.['regiao-imediata']?.['regiao-intermediaria']?.UF || item?.microrregiao?.mesorregiao?.UF || {};
  return { codigo: String(item.id || ''), nome: normalizeOfficialText(item.nome), uf: String(uf.sigla || '').toUpperCase() };
}).filter((item) => /^\d{7}$/.test(item.codigo) && item.nome && /^[A-Z]{2}$/.test(item.uf));

const byState = Object.fromEntries(states.map((state) => [
  state.sigla,
  municipalities.filter((item) => item.uf === state.sigla).map((item) => [item.codigo, item.nome])
]));

const metadata = {
  fonte: 'IBGE API de Localidades',
  fonteUrl: source,
  apiVersao: apiVersion,
  catalogoGeradoEm: catalogDate,
  quantidadePaises: countries.length,
  quantidadeUfs: states.length,
  quantidadeMunicipios: municipalities.length
};

const js = [
  '/** Arquivo gerado por scripts/update_locality_catalog.mjs. Nao editar manualmente. */',
  `var CORE_LOCALIDADES_CATALOGO_META = Object.freeze(${JSON.stringify(metadata)});`,
  `var CORE_LOCALIDADES_PAISES = Object.freeze(${JSON.stringify(countries)});`,
  `var CORE_LOCALIDADES_UFS = Object.freeze(${JSON.stringify(states)});`,
  `var CORE_LOCALIDADES_MUNICIPIOS_POR_UF = Object.freeze(${JSON.stringify(byState)});`,
  ''
].join('\n');

fs.writeFileSync(path.join(root, '39a_core_locality_catalog_data.js'), js, 'utf8');

const portalArg = process.argv.find((arg) => arg.startsWith('--portal-output='));
if (portalArg) {
  const target = path.resolve(root, portalArg.slice('--portal-output='.length));
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, `${JSON.stringify({ metadata, countries, states, municipalities }, null, 2)}\n`, 'utf8');
}

console.log(JSON.stringify(metadata));
