# Diagnostico de Registry V2 do Core

Este documento descreve o teste seguro para conferir se as keys de Pessoas V2 e
Vigencias V2 estao cadastradas no Registry da Planilha Geral.

A rotina e somente leitura:

- nao escreve em `MODULOS_STATUS`;
- nao escreve em `MODULOS_CONFIG`;
- nao abre nem altera as bases V2;
- nao cria triggers;
- nao altera PROD.

## Funcao publica

### `coreV2_runTesteResolverRegistryV2()`

Teste manual pronto para executar no editor do Apps Script.

Ele chama o diagnostico com `ambiente: 'DEV'` e retorna um item por key:

```javascript
{
  key: 'PESSOAS_V2_BASE',
  encontrada: true,
  spreadsheetId: '1abcde...WXYZ',
  sheetName: 'PESSOAS_BASE',
  ambiente: 'DEV',
  ativo: true,
  erro: ''
}
```

O `spreadsheetId` e sempre mascarado no retorno para evitar expor IDs internos
em logs ou prints de homologacao.

## Keys conferidas

Pessoas V2:

- `PESSOAS_V2_BASE`
- `PESSOAS_V2_IDENTIFICADORES`
- `PESSOAS_V2_MEMBROS_DETALHES`
- `PESSOAS_V2_VINCULOS_GEAPA`
- `PESSOAS_V2_MEMBROS_EVENTOS_VINCULO`
- `PESSOAS_V2_RESUMO_OPERACIONAL`

Vigencias V2:

- `VIGENCIAS_V2_SEMESTRES`
- `VIGENCIAS_V2_PERIODOS`
- `VIGENCIAS_V2_DIRETORIAS`
- `VIGENCIAS_V2_SEMESTRES_DIRETORIA`
- `VIGENCIAS_V2_CARGOS_CONFIG`
- `VIGENCIAS_V2_FUNCOES`
- `VIGENCIAS_V2_RESUMO_ATUAL`

## Por que este teste usa o Registry bruto

As APIs comuns do Registry, como `coreGetRegistry()` e `coreGetSheetByKey()`,
filtram por ambiente atual (`GEAPA_ENV`) e por `ATIVO=SIM`.

Isso e correto para execucao normal, mas pode confundir a homologacao:

- se o script estiver com `GEAPA_ENV=PROD`;
- e a key existir apenas em `AMBIENTE=DEV`;
- a key pode parecer ausente para o resolvedor filtrado.

Por isso `coreV2_runTesteResolverRegistryV2()` consulta o Registry bruto e
procura explicitamente as linhas `AMBIENTE=DEV`. O objetivo e separar:

- key realmente ausente;
- key existente em outro ambiente;
- key encontrada, mas inativa;
- key encontrada e ativa.

## Interpretacao do resultado

`ok=true` significa:

- todas as keys esperadas foram encontradas em `DEV`;
- todas estao ativas;
- o Registry bruto foi lido sem erro.

`ok=false` pode indicar:

- `KEY_NAO_ENCONTRADA`;
- `KEY_SEM_AMBIENTE_DEV`;
- `KEY_INATIVA`;
- erro estrutural ao ler o Registry.

## Ordem recomendada de homologacao

1. Rodar `coreClearRegistryCache()`.
2. Rodar `coreV2_runTesteResolverRegistryV2()`.
3. Confirmar que todas as keys retornam `encontrada=true` e `ativo=true`.
4. Rodar `coreV2_conferirConfiguracao({ ambiente: 'DEV' })`.
5. Rodar `coreV2DiagnosticoGeral({ dryRun: true })`.
6. Se ainda houver `keys indisponiveis`, conferir se o script publicado esta
   usando `GEAPA_ENV=DEV` ou se as rotinas precisam receber `ambiente: 'DEV'`
   explicitamente.

## Observacoes operacionais

- O teste nao valida se os cabecalhos das abas estao completos. Essa validacao
  continua nas rotinas `corePessoasV2Diagnostico`,
  `coreVigenciasV2Diagnostico` e `coreV2DiagnosticoGeral`.
- O teste nao cria aliases no Registry.
- O teste nao corrige linhas duplicadas.
- O teste nao muda o modo de `MODULOS_CONFIG`.
